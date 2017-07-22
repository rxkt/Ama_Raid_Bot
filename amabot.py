import discord
import datetime
import asyncio
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from discord.ext import commands



client = discord.Client()

scope = ['https://spreadsheets.google.com/feeds']

update_message = """
Hello, please update your answers to Amaterasu's 30 man form. IF YOU REMOVE A CHARACTER, PLEASE PM XPECT ON DISCORD.

IF YOU DO NOT EDIT YOUR FORM, YOUR DATA FROM THE PREVIOUS WEEK WILL BE USED.
"""

no_member_found = """Unable to find you in the spreadsheet. Please go to

<%s>

to sign up for raids.""" % "https://docs.google.com/forms/d/e/1FAIpQLSfr_P7wEhtq85H2mt7VO2_DETBhK-N-ihg7bBImcwsq88GOEQ/viewform"

class amabot(discord.Client):
    """Cog that scrapes Amaterasu's Raid spreadsheets!"""
    
    
    def __init__(self, bot):
        self.bot = bot

    def is_me(m):
        return m.author == client.user


    #When called, notifies all users who have not updated their spreadsheet within the past week to do so.
    @commands.command(pass_context=True)
    @client.event
    async def notify(self, ctx):
        """Mentions all people on the spreadsheet whose timestamps on their form responses are not updated for a period of time.""" 

        server = ctx.message.server
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        wsh = sh.worksheet("InfoParse")
        await self.bot.say("opened sheet . . .")
        
        memberIndex = 2
        memberList = list()
        todayDate = datetime.datetime.today()
        unmessaged_list = list()
        
        while True:
            s = wsh.cell(memberIndex,1).value
            name = wsh.cell(memberIndex, 3).value
            link = wsh.cell(memberIndex, 4).value
            
            if s == '':
                break
            s=s.split(' ')[0]
            
            if s!= '':
                try:
                    prevDate = datetime.datetime.strptime(s,"%m/%d/%Y")
                except:
                    if s=='#REF!':
                        await self.bot.say("#REF! found on line %d. Delete the row in the spreadsheet" % (memberIndex))
                    else:
                        await self.bot.say("Unknown error parsing timestamp.")
                        
                if abs((prevDate-todayDate).days) > 1: #exclude people who have updated yesterday?
                    
                    memberList.append( name)
                    member = server.get_member_named( wsh.cell(memberIndex,2).value.strip(' ') )
                    try:
                        if link != '':
                            await self.bot.send_message(member, "%s\n\n Your link is: %s\nIf you happen to have the wrong link, PM an officer on discord." % (update_message,link))
                        else:
                            await self.bot.send_message(member, "%s\n\n Your link is not available right now. PM an officer for assistance updating your sheets." % (update_message))
                        await self.bot.say("Messaged " + name )
                            
                    except:
                        await self.bot.say("Could not message " + name )
                        unmessaged_list.append( name )          
            memberIndex+=1


        if len(unmessaged_list) > 0:
            await self.bot.say(unmessaged_list)
        else:
            await self.bot.say("Everyone messaged.")


    #Sends a discord message to the user that called this function with their respective link to their google forms
    @commands.command(pass_context=True)
    @client.event
    async def link(self, ctx, *, user: discord.Member=None):
        "Retrieves the link to the form for signing up."

        author = ctx.message.author
        if not user:
            user = author

        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        wsh = sh.worksheet("InfoParse")
        #debug
        #await self.bot.say("opened sheet . . .")
        try:
            cell = wsh.find(str(user.id))
            link = wsh.cell(cell.row,4).value
            await self.bot.send_message(user, link)
            await self.bot.say("Sent you the link, check your DMs <:blobshiro:331164843417665539>")
        except:
            if 'Amaterasu' in [str(i) for i in user.roles]:
                await self.bot.send_message(user,no_member_found)
            await self.bot.say("You are not on the spreadsheet for Harrowhold raids.")

    @commands.command(pass_context=True)
    @client.event
    async def test(self, ctx, *, user: discord.Member=None):
        "Test function for members"
        author = ctx.message.author
        if not user:
            user = author
            
    #Sends a discord message to the user that called this function
    #with their respective performance ratings
    @commands.command(pass_context=True)
    @client.event
    async def pp(self, ctx, *, user:discord.Member=None):
        "Returns your performance rating of all your characters for raids in Phase 4."

        author = ctx.message.author
        if not user:
            user = author

        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        wsh = sh.worksheet("InfoParse")
        
        try:
            cell = wsh.find(str(user.id))
            name = wsh.cell(cell.row,3).value

        except:
            await self.bot.say("You are not on the spreadsheet for Harrowhold raids.")
            return

        await self.bot.say("Opened 'consolidated names'...")
        wsh_names = sh.worksheet("Consolidated Names")
        main_names = wsh_names.col_values(6)
        row_index = main_names.index(name)+1
        
        alt_names = [item for item in wsh_names.row_values(row_index) if item!= ''][1:]
        

        
        #deaths
        #number of attempts
        #pp, calculation
        #if no runs-> no runs yet! 

        
    

    #Assigns the discord IDs for each given discord tag on the spreadsheet
    @commands.command(pass_context=True)
    @client.event
    async def assign_ids(self, ctx, *, user: discord.Member=None):
        "Updates the ID of every discord member on the spreadsheet."

        server=ctx.message.server
        
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        wsh = sh.worksheet("InfoParse")
        await self.bot.say("opened sheet . . .")

        memberIndex=2
        while True:
            timestamp = wsh.cell(memberIndex,1).value
            discord_tag = wsh.cell(memberIndex, 2).value
            name = wsh.cell(memberIndex, 3).value
            if timestamp=='':
                break
            
            member = server.get_member_named( wsh.cell(memberIndex,2).value.strip(' ') )
            if member==None:
                await self.bot.say("Did not update ID for %s" % name)
            else:
                wsh.update_cell(memberIndex, 5, member.id)
                await self.bot.say("Updated ID for %s" % name)

            memberIndex+=1
        
        
        
    # Lists all raids not yet occurred. useless for now, may want to put this into a PM.
    @commands.command()
    async def raidtimes(self):
        "Lists all the raid that have not yet occurred."

        await self.bot.say("Searching the spreadsheets for raids...")
        
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        wsh = sh.worksheet("raid lineup")

        raidIndex = 1
        raidTimes=""
        
        while (wsh.cell(1,raidIndex).value != '' ):
            raidTimes+= "%d) %s\n " % ( (raidIndex/5+1) , wsh.cell(1,raidIndex).value )
            raidIndex+=5

        await self.bot.say("There are " + str( int(raidIndex/5) ) + " raids this week.\n" + raidTimes)

def setup(bot):
    bot.add_cog(amabot(bot))
