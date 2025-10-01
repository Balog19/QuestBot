import os
from typing import Union, Optional

import discord
from discord.ext import commands

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import get_effective_format, format_cell_range


# =========================
# Google Sheets setup
# =========================
SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]
CREDS = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", SCOPE)
GCLIENT = gspread.authorize(CREDS)
SHEET = GCLIENT.open("Tracking Test").worksheet("Points Tracking")

# =========================
# Discord bot setup
# =========================
intents = discord.Intents.default()
intents.message_content = True
intents.members = True  # needed so mentions resolve properly
bot = commands.Bot(command_prefix="!", intents=intents)

REQUIRED_ROLE_NAME = "Staff"

# Columns we increment via commands
POINT_COLUMNS = ["1st Places", "2nd Places", "3rd Places", "Quest", "Bonus", "Participation"]
POINTS_TOTAL_HEADER = "Points (X Edit)"


# -------------------------
# Helpers
# -------------------------
def col_to_letter(col: int) -> str:
    result = ""
    while col > 0:
        col, r = divmod(col - 1, 26)
        result = chr(65 + r) + result
    return result

def header_indexes():
    headers = SHEET.row_values(1)
    return {h.strip().lower(): i + 1 for i, h in enumerate(headers)}

def get_col_index(header_name: str) -> int:
    idx = header_indexes().get(header_name.strip().lower())
    if not idx:
        raise ValueError(f"Column '{header_name}' not found.")
    return idx

def get_names(member: discord.Member):
    if isinstance(member, discord.Member) and member.nick:
        return member.nick, member.name
    return member.name, member.name

def build_points_formula(row: int) -> str:
    return f"=SUMPRODUCT(E{row}:J{row}, TRANSPOSE($P$2:$P$7)) + K{row} - L{row}"

def increment_many(members, column_header: str, delta: int = 1):
    """
    Add or subtract 'delta' points for each member in the given column.
    If member row doesn't exist and delta > 0, a row is created with formatting + totals formula.
    """
    nick_col = get_col_index("Discord Nickname")
    target_col = get_col_index(column_header)

    nickname_list = SHEET.col_values(nick_col)
    target_list = SHEET.col_values(target_col)
    if len(target_list) < len(nickname_list):
        target_list += [""] * (len(nickname_list) - len(target_list))

    updates = []

    for member in members:
        nickname, username = get_names(member)

        if nickname in nickname_list:
            row_idx = nickname_list.index(nickname) + 1
            cell_val = target_list[row_idx - 1]
            try:
                current = int(cell_val)
            except (TypeError, ValueError):
                current = 0
            new_val = current + delta
            if new_val < 0:
                new_val = 0  # clamp at zero
            updates.append({
                "range": f"{col_to_letter(target_col)}{row_idx}",
                "values": [[new_val]]
            })
        else:
            # Only create a new row when adding (delta > 0)
            if delta > 0:
                new_row = len(nickname_list) + 1
                body = {
                    "requests": [{
                        "insertDimension": {
                            "range": {
                                "sheetId": SHEET.id,
                                "dimension": "ROWS",
                                "startIndex": new_row - 1,
                                "endIndex": new_row
                            },
                            "inheritFromBefore": True
                        }
                    }]
                }
                SHEET.spreadsheet.batch_update(body)

                # Nick + Username
                updates.append({
                    "range": f"{col_to_letter(nick_col)}{new_row}",
                    "values": [[nickname]]
                })
                updates.append({
                    "range": f"{col_to_letter(get_col_index('Discord Username'))}{new_row}",
                    "values": [[username]]
                })

                # Target column value = delta
                updates.append({
                    "range": f"{col_to_letter(target_col)}{new_row}",
                    "values": [[delta]]
                })

                # Reset other point columns to 0 for the new row
                for header in POINT_COLUMNS:
                    if header != column_header:
                        try:
                            c = get_col_index(header)
                            updates.append({"range": f"{col_to_letter(c)}{new_row}", "values": [[0]]})
                        except Exception:
                            pass

                # Add totals formula
                points_col = get_col_index(POINTS_TOTAL_HEADER)
                updates.append({
                    "range": f"{col_to_letter(points_col)}{new_row}",
                    "values": [[build_points_formula(new_row)]]
                })

    if updates:
        SHEET.batch_update(updates, value_input_option="USER_ENTERED")


# -------------------------
# Bot events / commands
# -------------------------
@bot.event
async def on_ready():
    print(f"Bot is ready. Logged in as {bot.user}")

def staff_only():
    async def predicate(ctx):
        return any(r.name == REQUIRED_ROLE_NAME for r in getattr(ctx.author, "roles", []))
    return commands.check(predicate)


# ===== Commands with only ±1 (remove optional) =====
def make_simple_command(name, header):
    @bot.command(name=name)
    @staff_only()
    async def cmd(ctx, first: Union[discord.Member, str, None] = None, *rest: discord.Member):
        """
        Usage:
          !first @User                 -> +1
          !first remove @User @User2   -> -1 each
        """
        members = []
        delta = 1

        if isinstance(first, str) and first.lower() == "remove":
            delta = -1
            members = list(rest)
        else:
            if isinstance(first, discord.Member):
                members = [first] + list(rest)
            else:
                members = list(rest)

        if not members:
            await ctx.send("No members provided.")
            return

        increment_many(members, header, delta)
        names = ", ".join(m.display_name for m in members)
        if delta > 0:
            await ctx.send(f"✅ Added 1 point to {header} for: {names}")
        else:
            await ctx.send(f"✅ Removed 1 point from {header} for: {names}")

make_simple_command("first", "1st Places")
make_simple_command("second", "2nd Places")
make_simple_command("third", "3rd Places")
make_simple_command("quest", "Quest")
make_simple_command("participation", "Participation")


# ===== BONUS: supports numbers and remove numbers =====
@bot.command(name="bonus")
@staff_only()
async def bonus_cmd(ctx, *args):
    """
    Patterns supported:
      !bonus @User                          -> +1
      !bonus 5 @User1 @User2               -> +5 each
      !bonus remove @User                  -> -1
      !bonus remove 3 @User1 @User2        -> -3 each
    """
    if not args:
        await ctx.send("No members provided.")
        return

    sign = +1
    amount = 1
    i = 0

    # Handle "remove"
    if args[0].lower() == "remove":
        sign = -1
        i += 1

    # Handle number after (optional)
    if i < len(args) and args[i].isdigit():
        amount = int(args[i])
        i += 1

    # Get all mentioned members directly
    members = ctx.message.mentions

    if not members:
        await ctx.send("No members provided.")
        return

    delta = sign * amount
    increment_many(members, "Bonus", delta)

    names = ", ".join(m.display_name for m in members)
    if delta > 0:
        await ctx.send(f"✅ Added {delta} point(s) to Bonus for: {names}")
    else:
        await ctx.send(f"✅ Removed {abs(delta)} point(s) from Bonus for: {names}")


# ===== POINTS: check your own score (no role required) =====
@bot.command(name="points")
async def points_cmd(ctx):
    """
    Anyone can use !points to check their own score.
    """
    member = ctx.author
    nickname, username = get_names(member)

    # Column indexes
    nick_col = get_col_index("Discord Nickname")
    points_col = get_col_index(POINTS_TOTAL_HEADER)

    # All nicknames in the sheet
    nickname_list = SHEET.col_values(nick_col)

    if nickname in nickname_list:
        row_idx = nickname_list.index(nickname) + 1
        points_val = SHEET.cell(row_idx, points_col).value
        try:
            points_val = int(points_val)
        except (TypeError, ValueError):
            points_val = 0
    else:
        points_val = 0

    await ctx.send(f"{member.display_name}, you currently have **{points_val}** points.")


# =========================
# Run the bot
# =========================
bot.run("DISCORD_BOT_TOKEN")
