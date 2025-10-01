import os
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
bot = commands.Bot(command_prefix="!", intents=intents)

REQUIRED_ROLE_NAME = "Staff"

# Columns we increment via commands
POINT_COLUMNS = ["1st Places", "2nd Places", "3rd Places", "Quest", "Bonus", "Participation"]
POINTS_TOTAL_HEADER = "Points (X Edit)"  # the computed total column


# -------------------------
# Helpers
# -------------------------
def col_to_letter(col: int) -> str:
    """1 -> A, 27 -> AA, etc."""
    result = ""
    while col > 0:
        col, r = divmod(col - 1, 26)
        result = chr(65 + r) + result
    return result

def header_indexes():
    headers = SHEET.row_values(1)
    return {h.strip().lower(): i+1 for i, h in enumerate(headers)}

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
    Add points for a set of members in one column.
    delta = +1 always (only adding for now).
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
            updates.append({
                "range": f"{col_to_letter(target_col)}{row_idx}",
                "values": [[new_val]]
            })
        else:
            # If user doesn’t exist yet, create row
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

            updates.append({
                "range": f"{col_to_letter(nick_col)}{new_row}",
                "values": [[nickname]]
            })
            updates.append({
                "range": f"{col_to_letter(get_col_index('Discord Username'))}{new_row}",
                "values": [[username]]
            })
            updates.append({
                "range": f"{col_to_letter(target_col)}{new_row}",
                "values": [[1]]
            })
            # Reset other point cols
            for header in POINT_COLUMNS:
                if header != column_header:
                    try:
                        c = get_col_index(header)
                        updates.append({"range": f"{col_to_letter(c)}{new_row}", "values": [[0]]})
                    except:
                        pass
            # Add formula
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

def make_point_command(name, header):
    @bot.command(name=name)
    @staff_only()
    async def cmd(ctx, *members: discord.Member):
        if not members:
            await ctx.send("No members provided.")
            return

        increment_many(members, header, 1)
        await ctx.send(f"✅ Added 1 point to {header}.")

# Create all point commands
make_point_command("first", "1st Places")
make_point_command("second", "2nd Places")
make_point_command("third", "3rd Places")
make_point_command("quest", "Quest")
make_point_command("bonus", "Bonus")
make_point_command("participation", "Participation")


# =========================
# Run the bot
# =========================
bot.run("DISCORD_BOT_TOKEN")
