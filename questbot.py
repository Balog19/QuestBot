import discord
import os
from discord.ext import commands
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ---------- Google Sheets setup ----------
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]
creds = ServiceAccountCredentials.from_json_keyfile_name(
    "credentials.json", scope
)
gclient = gspread.authorize(creds)
sheet = gclient.open("TESTPOINTS").worksheet("Points Tracking")

# ---------- Discord bot setup ----------
intents = discord.Intents.default()
intents.message_content = True
intents.reactions = True
intents.messages = True
bot = commands.Bot(command_prefix="!", intents=intents)

TARGET_CHANNEL_NAME = "quests"
REQUIRED_ROLE_NAME = "Staff"

@bot.event
async def on_ready():
    print(f"Bot is ready. Logged in as {bot.user}")

@bot.event
async def on_raw_reaction_add(payload):
    guild = bot.get_guild(payload.guild_id)
    channel = guild.get_channel(payload.channel_id)

    guild = bot.get_guild(payload.guild_id)
    channel = guild.get_channel(payload.channel_id)
    if channel is None:
        return

    if channel.name != TARGET_CHANNEL_NAME:
        return

    member = await guild.fetch_member(payload.user_id)
    if member is None:
        return

    has_required_role = any(role.name == REQUIRED_ROLE_NAME for role in member.roles)
    if not has_required_role:
        return

    # Fetch the message that was reacted to
    message = await channel.fetch_message(payload.message_id)
    original_author = message.author.name


    try:
        column_d = sheet.col_values(4)  # Column D (1‑indexed)

        if original_author in column_d:
            # Row already exists – increment column I
            row_index = column_d.index(original_author) + 1
            current_value = sheet.cell(row_index, 9).value  # Column I
            try:
                current_number = int(current_value)
            except (TypeError, ValueError):
                current_number = 0
            sheet.update_cell(row_index, 9, current_number + 1)
        else:
            # New entry – append at the end
            new_row_index = len(column_d) + 1
            sheet.update_cell(new_row_index, 4, original_author)  # Column D
            sheet.update_cell(new_row_index, 9, 1)                # Column I

        await channel.send(f"Added quest point for **{original_author}**.")
    except Exception as e:
        print(f"Error interacting with sheet: {e}")

# ---------- Run the bot ----------
bot.run(os.getenv("DISCORD_BOT_TOKEN"))
