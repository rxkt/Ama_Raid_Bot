import discord

import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds']

credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)

gc = gspread.authorize(credentials)

#formula: wsh.acell('b1').input_value

from discord.ext import commands



class amabot:
    """Cog that scrapes Amaterasu's Raid spreadsheets!"""
    
    
    def __init__(self, bot):
        self.bot = bot

    @commands.command(pass_context=True)
    async def callsnow(self, ctx):
        """ test"""
        #"""Displays the given raid number in full, with all roles."""
        server = ctx.message.server
        member = server.get_member_named("Snow#4465")
        await self.bot.say(member.mention)

    @commands.command()
    async def raidtimes(self):
        "Lists all the raid times that have not yet occurred."
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
