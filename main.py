import discord
import os
from discord.ext import commands
from dotenv import load_dotenv
import pandas as pd
import openpyxl

intents = discord.Intents.default()
intents.members = True
intents.message_content = True
intents.typing = False
intents.presences = False

load_dotenv()
TOKEN = os.getenv('DISCORD_TOKEN')
GUILD = os.getenv('DISCORD_GUILD')
SERVER_ID = 1068333068668125236

print(GUILD)
client = discord.Client(intents=intents)

bot = commands.Bot(command_prefix="!", intents=intents)

registered = pd.read_excel('CxC_signups.xlsx', sheet_name=0)

student_nums = registered['Student Number'].tolist()
checkedin = []


async def dm_about_roles(member):
    print(f"DMing {member.name}...")

    await member.send(
        f"""Hi {member.name}, welcome to {member.guild.name}!

   What is your student number?

    """
    )


@bot.event
async def on_member_join(member):
    await dm_about_roles(member)


async def assign_roles(message):

    num = message.content
    found = int(num) in student_nums
    verified = int(num) in checkedin
    if found and not verified:
        server = bot.get_guild(SERVER_ID)

        role = discord.utils.get(server.roles, name='Participant')
        authid = message.author.id
        member = await server.fetch_member(authid)

        try:
            await member.add_roles(role)
        except Exception as e:
            print(e)
            await message.channel.send("Error assigning roles.")
        else:
            await message.channel.send(f"""You've been verified. Welcome to CXC! """)
            checkedin.append(int(num))
    else:
        await message.author.send("I wasn't able to find your name in the registered list or you've already checked "
                                  "in to the event")


@bot.event
async def on_message(message):
    print("Saw a message...")

    if message.author == bot.user:
        return  # prevent responding to self

    if isinstance(message.channel, discord.channel.DMChannel):
        await assign_roles(message)
        return

    if message.content.startswith("!verify"):
        await dm_about_roles(message.author)


bot.run(TOKEN)
