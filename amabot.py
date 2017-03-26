import discord
import datetime
import asyncio

import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds']


from discord.ext import commands

client = discord.Client()


update_message = """
Hello, please update your answers to Amaterasu's 30 man form. You can change your hours available or edit your characters. IF YOU REMOVE A CHARACTER, PLEASE PM XPECT ON DISCORD.

If nothing changed, you still need to submit the form again (click SUBMIT on the form to update the timestamp).

If you already updated your form for next week, ignore this message.

P.S
Don't change your discord ID. The bot wil not PM you, your schedule can be messed up, and you risk getting yellow/orange cards. Change your nickname instead."""

class amabot(discord.Client):
    """Cog that scrapes Amaterasu's Raid spreadsheets!"""
    
    
    def __init__(self, bot):
        self.bot = bot

    def is_me(m):
        return m.author == client.user

    @commands.command(pass_context=True)
    @client.event
    async def notify(self, ctx):
        """Mentions all people on the spreadsheet whose timestamps on their form responses are not updated for a period of time."""

        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("Raid info spreadsheet")
        wsh = sh.worksheet("Info Parse")
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
                    server = ctx.message.server
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

    #test function.
    @commands.command(pass_context=True)
    @client.event
    async def pm(self,ctx,user=discord.Member):
        "Test function"
        server = ctx.message.server
        #test using IDs instead of name#discrim
        member = server.get_member_named(user)
        await self.bot.send_message(member.id, update_message)

        
    #assigns the role given.
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
        
    #useless for now, may want to put this into a PM. Lists all raids not yet occurred.
    @commands.command()
    async def raidtimes(self):
        "Lists all the raid that have not yet occurred."
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("Raid info spreadsheet")
        wsh = sh.worksheet("raid_lineup2")

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
