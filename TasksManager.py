import discord
from discord.ext import commands
from discord import app_commands
import requests, sqlite3, json, datetime, openpyxl, asyncio, random

prefix = '-'
bot = commands.Bot(command_prefix=prefix, intents=discord.Intents.all())  #Discord bot
botAuthorID = 503484838029033483
configFile = 'config.json'
languagePackageFile = 'surface_language.json'

#Files checking -----------------------------------------------------------------------------------

#Check json config file
def LoadConfigFile():
    global configFile
    global languagePackageFile
    isConfigExists = False  #check if config file exists or not

    try:
        with open(configFile, 'r', encoding="utf-8") as configFile:
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
        with open(configFile, 'w', encoding="utf-8") as newConfigFile:
            modifyTime = str(datetime.datetime.now())
            configInit = {
                "CreatedTime" : modifyTime,
                "LastModifiedTime" : modifyTime
            }

            #load language package
            with open(languagePackageFile, 'r') as langPackage:
                data = json.load(langPackage)
                configInit["Language"] = data[0]["Language"]

            json.dump(configInit, newConfigFile, indent=2, ensure_ascii=False)
            print("Created config file.")

            return configInit    
        
print("------ Checking config file ------")
botConfig = LoadConfigFile()  #setting config data from json file
print()

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

print("------ Checking excel workbooks ------")
#invoke the function for checking the existence of excel workbook everytime while launching
sheetTitle = LoadScheduleList()
print()

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
            await asyncio.sleep(1)
    except Exception as ex:
        print(ex)

'''
@bot.event
async def on_raw_reaction_add(reaction : discord.RawReactionActionEvent):
    channel = bot.get_channel(reaction.channel_id)                       
    await channel.send(reaction.message_id)
'''

@bot.event
async def on_message(msg):  #When bot listened specific message from user
    if msg.author.bot:
        return
    
    if "114" in msg.content:
        await msg.reply("114514", mention_author=False)

    if "一個" in msg.content:
        await msg.reply("你是一個 一個一個 哼哼哼啊啊啊啊啊啊", mention_author=False)

#Bot slash commands -------------------------------------------------------------------------------
        
@bot.tree.command(name="info", description="Bot introduction.")  #Enter the name to use this command function
async def Introduce(interaction : discord.Interaction):
    botAuthor = await bot.fetch_user(botAuthorID)  #fetch the author

    with open(configFile.name, 'r', encoding='utf-8') as config:
        configData = json.load(config)

    #create main embed introduction
    infoEmbed = discord.Embed(
        title=bot.user.name,
        url='https://youtu.be/dQw4w9WgXcQ',
        description="I'm a schedule manager bot.",
        color=discord.Color.random(),
        timestamp=datetime.datetime.now()
    )
    #set other decorations
    infoEmbed.set_author(name=f"Author : {botAuthor.name}")
    infoEmbed.set_thumbnail(url=botAuthor.avatar.url)
    infoEmbed.set_footer(text="114514")

    infoEmbed.add_field(name="Language :", value=configData['Language'])

    await interaction.response.send_message(embed=infoEmbed, ephemeral=False)
    #If ephemeral = True, that means only you can see the reply.

@bot.tree.command(name="reply")
@app_commands.describe(ctx="Enter any text : ")  #description for each variable
async def reply(interaction : discord.Interaction, ctx : str):
    await interaction.response.send_message(f'Hi, {interaction.user.name}. You just entered `{ctx}`.', ephemeral=True)

@bot.tree.command(name="taskslist", description="Watch the to-do list in your schedule")  #Show the to-do list
async def tasksList(interaction : discord.Interaction): 
    todoListEmbed = discord.Embed(  #set the contents in the embed
        title="Your Schedule :",
        url='https://youtu.be/dQw4w9WgXcQ',
        description="To-Do List",
        color=discord.Color.random(),
        timestamp=datetime.datetime.now()
    )
    todoListEmbed.set_thumbnail(url=bot.user.avatar.url)
    todoListEmbed.set_footer(text="schedule")

    todoListEmbed.add_field(name="Tasks", value="1145141919810", inline=False)

    await interaction.response.send_message(embed=todoListEmbed, ephemeral=False)  #send the embed message
    msg = await interaction.original_response()  #get the first message which sent by this slash command

    #when there's a reaction to the embed message, edit it.
    while True:
        #to check if the user and the message is correct,
        #if the user who added the reaction is the same as the one who used this command,
        #and if the target of the reaction is the embed sent by this command:
        def checkRection(reaction, user):  
                return user == interaction.user and reaction.message.id == msg.id

        reaction, user = await bot.wait_for(  #listen the reaction
            'reaction_add',
            timeout = None,  #this listening action doesn't expire
            check = checkRection  #invoke the checking function
            )

        #edit original embed message
        todoListEmbed.add_field(name=str(len(todoListEmbed.fields)), value=str(msg.id), inline=False)
        await interaction.edit_original_response(embed=todoListEmbed)

#Config commands --------------------------------
@bot.tree.command(name="config", description="Check the bot config.")
async def checkConfig(interaction : discord.Interaction):
    with open(configFile.name, 'r', encoding='utf-8') as config:
        configData = json.load(config)
    await interaction.response.send_message(f"Bot config :\n{configData}", ephemeral=False)
 
@bot.tree.command(name="changelanguage", description="Change your surface language.")
@app_commands.describe(language="Type en/ch/jp/kr, en:English, ch:Chinese, jp:Japanese, kr:Korean.")
async def changeLanguage(interaction : discord.Interaction, language : str):
    global configFile
    global languagePackageFile

    isNotChanged = False
    with open(configFile.name, 'r', encoding="utf-8") as config:
        configData = json.load(config)

    with open(configFile.name, 'w', encoding="utf-8") as config:
        with open(languagePackageFile, 'r') as langPackage:
            data = json.load(langPackage)
            match language:
                case 'en':
                    configData["Language"] = data[0]["Language"]
                case 'ch':
                    configData["Language"] = data[1]["Language"]
                case 'jp':
                    configData["Language"] = data[2]["Language"]
                case 'kr':
                    configData["Language"] = data[3]["Language"]
                case _:
                    isNotChanged = True  #language isn't changed
                    await interaction.response.send_message("Error, enter again.", ephemeral=True)
            
            if not isNotChanged:
                configData['LastModifiedTime'] = str(datetime.datetime.now())
                print(f"Modified config file at {configData['LastModifiedTime']}")
            
            json.dump(configData, config, indent=2, ensure_ascii=False)
            
    if not isNotChanged:  #if language changed
        await interaction.response.send_message(f"Surface language is {configData['Language']} now.", ephemeral=False)
        print(f"Change language into {configData['Language']}")

@bot.tree.command(name="setyourtimeunit", description="Set time unit to measure the total time each task needs.")
@app_commands.describe(time="If you type '30', it means it takes 30/60/90... MINUTES to finish your each task")
async def setTimeUnit(interaction : discord.Interaction, time : int):
    with open(configFile.name, 'r', encoding='utf-8') as config:
        configData = json.load(config)

    try:
        with open(configFile.name, 'w', encoding='utf-8') as config:
            configData['TimeUnit'] = time
            configData['LastModifiedTime'] = str(datetime.datetime.now())
            json.dump(configData, config, indent=2, ensure_ascii=False)

            print(f"Modified config file at {configData['LastModifiedTime']}")
            print(f"Changed time unit into {configData['TimeUnit']}")

            await interaction.response.send_message(f"Modified time unit to {configData['TimeUnit']}", ephemeral=False)
    except Exception as ex:
        print(ex)
        await interaction.response.send_message("Error, please enter digits!", ephemeral=True)

#Excel commands ---------------------------------
@bot.tree.command(name="todolist", description="test command to see the title of worksheet.")
async def checkScheduleList(interaction : discord.Interaction):
    await interaction.response.send_message(f"Worksheet title :\n{sheetTitle}", ephemeral=True)

#--------------------------------------------------------------------------------------------------

bot.run(json.load(open('token.json', 'r'))['token'])  #run this bot