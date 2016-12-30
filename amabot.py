import discord
import datetime
import asyncio

import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds']



#formula: wsh.acell('b1').input_value

from discord.ext import commands

client = discord.Client()
update_message = """
Hello, please update your answers to Amaterasu's 30 man form. You can change your hours available, or add an alt that can do the raid.

If nothing changed, please check your answers and click SUBMIT on the form to update the timestamp.

You will continue to receive these notifications until you update your answers to the form. If you need the link to your answers, PM Kyangi on discord.

"""

class amabot(discord.Client):
    """Cog that scrapes Amaterasu's Raid spreadsheets!"""
    
    
    def __init__(self, bot):
        self.bot = bot
        
    @commands.command(pass_context=True)
    @client.event
    async def mentionall(self, ctx):
        """Mentions all people on the spreadsheet whose timestamps on their form responses are not updated for a week."""

        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        #open sheet 1, with form responses.
        #loop thru all members, check if timestamp is greater than a week ago.
        #send em all messages or mention them? messages are better, will change soon.
        sh = gc.open("Raid info spreadsheet")
        wsh = sh.worksheet("Info Parse")

        
        memberIndex = 2
        memberList = list();
        #loop through all members. if they have not updated their thing, PM them. ( we will add them to list for now )
        while True:
            #we want to avoid calling the spreadsheet data twice because the bot is slow enough already...
            s = wsh.cell(memberIndex,1).value
            if s == '':
                break
            s=s.split(' ')[0]
            if s!= '':
                #get the last updated date, and today
                prevDate = datetime.datetime.strptime(s,"%m/%d/%Y")
                todayDate = datetime.datetime.today()
                #if the difference in days is greater than 6, add em to the list.
                if abs((prevDate-todayDate).days) > 6:
                    memberList.append( wsh.cell(memberIndex,3).value )

                    ''' uncomment as soon as we need this running
                    server = ctx.message.server
                    member = server.get_member_named( wsh.cell(memberIndex,2).value )
                    await self.bot.send_message(member, update_message)
                    '''
            memberIndex+=1

        '''debug, messages a member once
        server = ctx.message.server
        member = server.get_member_named("rxkt#2283")
        await self.bot.send_message(member,update_message)
        '''
        
        await self.bot.say(memberList)

    @commands.command(pass_context=True)
    @client.event
    async def pm_kyang(self,ctx):
        "Test function"
        server = ctx.message.server
        member = server.get_member_named("kyangi")
        await self.bot.send_message(member, "hi this is bot")

        
    #assigns the role given.
    #adding this because the other bot may be unreliable
    #case sensitive for now, will update algorithm to get all roles instead of using dynamic array
    @commands.command(pass_context=True)
    async def iam(self, ctx: commands.Context, role: discord.Role ):
        "List of possible roles :\n DPS, Healer, Tank, Priest, Mystic, Overwatch, ERP, League, Weeb. \n\n CASE SENSITIVE."
        author = ctx.message.author
        roles = [ 'DPS' , 'Healer', 'Tank' , 'Priest' , 'Mystic' , 'Overwatch' , 'ERP' , 'League' , 'Weeb' ]
        for x in roles:
            #.lower() is unneeded... but keeping for updated algo
            if str(role).lower() == x.lower():
                await self.bot.add_roles(author, role)
                await self.bot.say(":ok: You now have the " + x + " role.")
                return
        
        await self.bot.say("Not a valid role. ")
        

    @commands.command()
    async def raidtimes(self):
        "Lists all the raid times that have not yet occurred."
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("Raid info spreadsheet")
        wsh = sh.worksheet("raid_lineup")

        raidIndex = 1
        raidList = list();
        
        #while there is still a raid, create a dict and add it to the list
        while (wsh.cell(1,raidIndex).value != '' ):
            raidDict = {}
            raidDict['day']=wsh.cell(1,raidIndex).value
            raidDict['date']=wsh.cell(1,raidIndex+1).value
            raidDict['time']=wsh.cell(1,raidIndex+2).value
            #add all tanks, dps, priest, mys as arrays to keys
            
            raidList.append(raidDict)
            raidIndex+=5
        raidTimes=""

        for each in raidList:
            raidTimes+= str(raidList.index(each)+1) + ")\t" + each.get('date') + "\t@ " + each.get('time') + "\t(" + each.get('day') + ")\n"
        
        await self.bot.say("There are " + str( int(raidIndex/5) ) + " raids this week.\n" + raidTimes)

def setup(bot):
    bot.add_cog(amabot(bot))
