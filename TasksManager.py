import discord
from discord.ext import commands
from discord import app_commands
import requests, sqlite3, json, datetime, openpyxl, asyncio, random

prefix = '-'
bot = commands.Bot(command_prefix=prefix, intents=discord.Intents.all())  #Discord bot
botAuthorID = 503484838029033483

#Files checking -----------------------------------------------------------------------------------

#Check json config file
def LoadConfigFile():
    fileName = "config.json"
    isConfigExists = False  #check if config file exists or not
    try:
        with open(fileName, 'r', encoding="utf-8") as configFile:
            configData = json.load(configFile)

            #if except situation doesn't appear
            isConfigExists = True
            print("...Found config.json")
            print(configData)
            if isConfigExists == True:
                return configData
    except Exception as ex:
        print(ex)

    if isConfigExists == False:
        print("Config file is not found. Ready to create one.")
        with open(fileName, 'w', encoding="utf-8") as newConfigFile:
            modifyTime = str(datetime.datetime.now())
            configInit = {
                "CreatedTime" : modifyTime
            }
            configInit["LastModifiedTime"] = modifyTime

            json.dump(configInit, newConfigFile, indent=2, ensure_ascii=False)
            print("Created config file.")

            return configInit    

botConfig = LoadConfigFile()  #setting config data from json file

#Check schedule list in excel workbook
def LoadScheduleList():
    fileName = "To-Do List.xlsx"
    isExcelListExists = False
    try:
        workbook = openpyxl.load_workbook(fileName)
        isExcelListExists = True
        print("...Found the workbook of schedule list.")
        sheet = workbook.worksheets[0]

        print(sheet['A1'])
        return sheet.title
    except Exception as ex:
        print(ex)

    if isExcelListExists == False:
        print("Didn't find the Excel workbook. Ready to create one.")
        newWorkbook = openpyxl.Workbook()  #create new excel workbook
        sheet = newWorkbook.worksheets[0]  #first worksheet in the workbook
        sheet.title = "To-Do List"
        print("Created a workbook and set the title of the first worksheet.")

        sheet['A1'] = sheet.title
        newWorkbook.save(fileName)
        print("Saved the workbook.")

        return sheet.title

#invoke the function for checking the existence of excel workbook everytime while launching
sheetTitle = LoadScheduleList()

#Bot events ---------------------------------------------------------------------------------------

@bot.event
async def on_ready():  #When bot launched
    botAuthor = await bot.fetch_user(botAuthorID)  #fetch the author of this bot
    print(f"Bot Author : {botAuthor.name}")  #print the name of the author

    print(f'{bot.user} launched.')
    status = ['蔚藍檔案', 'Blue Archive', 'ブルーアーカイブ', '블루 아카이브']
    try:
        slashSync = await bot.tree.sync()
        print(f'Bot has {len(slashSync)} slash command(s)')
        while True:
            num = random.randint(0, len(status)-1)  #pick a status index randomly
            await bot.change_presence(status=discord.Status.online, activity=discord.Game(name=status[num]))
            await asyncio.sleep(5)
    except Exception as ex:
        print(ex)

@bot.event
async def on_message(msg):  #When bot listened specific message from user
    if msg.author.bot:
        return
    
    if "114" in msg.content:
        await msg.reply("114514", mention_author=False)

#Bot slash commands -------------------------------------------------------------------------------

@bot.tree.command(name="info")  #Enter the name to use this command function
async def Introduce(interaction : discord.Interaction):
    botAuthor = await bot.fetch_user(botAuthorID)  #fetch the author

    #create main embed introduction
    infoEmbed = discord.Embed(
        title=bot.user.name,
        url='https://youtu.be/dQw4w9WgXcQ',
        description="I'm a schedule manager bot.",
        color=discord.Color.random(),
        timestamp=datetime.datetime.now()
    )
    #set other decorations
    infoEmbed.set_author(name=botAuthor.name)
    infoEmbed.set_thumbnail(url=botAuthor.avatar.url)
    infoEmbed.set_footer(text="114514")

    await interaction.response.send_message(embed=infoEmbed, ephemeral=True)
    #If ephemeral = True, that means only you can see the reply.

@bot.tree.command(name="reply")
@app_commands.describe(ctx="Enter any text : ")  #description for each variable
async def reply(interaction : discord.Interaction, ctx : str):
    await interaction.response.send_message(f'Hi, {interaction.user.name}. You just entered `{ctx}`.', ephemeral=True)

@bot.tree.command(name="config", description="Check the bot config.")
async def checkConfig(interaction : discord.Interaction):
    await interaction.response.send_message(f"Bot config :\n{botConfig}", ephemeral=True)

@bot.tree.command(name="todolist", description="test command to see the title of worksheet.")
async def checkScheduleList(interaction : discord.Interaction):
    await interaction.response.send_message(f"Worksheet title :\n{sheetTitle}", ephemeral=True)

#--------------------------------------------------------------------------------------------------

bot.run(json.load(open('token.json', 'r'))['token'])  #run this bot