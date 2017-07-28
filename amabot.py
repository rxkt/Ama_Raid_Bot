import discord
import datetime
import asyncio
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from discord.ext import commands


scope = ['https://spreadsheets.google.com/feeds']

update_message = """
Hello, please update your answers to Amaterasu's 20 man form. IF YOU REMOVE/RETIRE A CHARACTER, PLEASE PM XPECT ON DISCORD.

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

    #to-do: server.get_member(userID)
    #When called, notifies all users who have not updated their spreadsheet within the past week to do so.
    @commands.command(pass_context=True)
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

                #script runs on saturday, new raid week on tuesday. 4 days sounds good
                if abs((prevDate-todayDate).days) > 3: 
                    
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
    async def link(self, ctx):
        "Retrieves the link to the form for signing up."

        author = ctx.message.author


        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        wsh = sh.worksheet("InfoParse")
        #debug
        #await self.bot.say("opened sheet . . .")
        try:
            cell = wsh.find(str(author.id))
            link = wsh.cell(cell.row,4).value
            await self.bot.send_message(author, link)
            await self.bot.say("Sent you the link, check your DMs <:blobshiro:331164843417665539>")
        except:
            if 'Amaterasu' in [str(i) for i in author.roles]:
                await self.bot.send_message(author,no_member_found)
            await self.bot.say("You are not on the spreadsheet for Harrowhold raids.")

    @commands.command(pass_context=True)
    async def test(self, ctx, *, user: discord.Member=None):
        "Test function for members"
        author = ctx.message.author
        if not user:
            user = author
        #someone else can be used as the parameter to bot call here
            
    #Sends a discord message to the user that called this function
    #with their respective performance ratings
    @commands.command(pass_context=True)
    async def pp(self, ctx):
        "Returns your performance rating of all your characters for raids in Phase 4."

        author = ctx.message.author

        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        wsh = sh.worksheet("InfoParse")
        
        try:
            cell = wsh.find(str(author.id))
            name = wsh.cell(cell.row,3).value

        except:
            await self.bot.say("You are not on the spreadsheet for Harrowhold raids.")
            return

        await self.bot.say("Finding your character names...")
        wsh_names = sh.worksheet("Consolidated Names")
        main_names = wsh_names.col_values(6)
        row_index = main_names.index(name)+1
        
        names_in_row = [item for item in wsh_names.row_values(row_index) if item!= '']
        main_index = names_in_row.index(name)
        alt_names = names_in_row[main_index:]

        await self.bot.say("Searching your statistics...")
        wsh_pp = sh.worksheet("PP Calculation2")
        pp_list = []
        runs=0
        #to solve multiple occurences in spreadsheet due to unsorted items, wsh.find was not used
        search_list = wsh_pp.col_values(2)[1:] #trim off the table header
        
        
        for each in alt_names:
            charpp_list = []
            
            #to solve multiple occurences in spreadsheet
            try:
                row_number = search_list.index(each)+2
            
                charpp_list.append(each)
                for x in range(4,7):
                    charpp_list.append(wsh_pp.cell(row_number,x).value )
                pp_list.append(charpp_list)
                
                try:
                    runs+=float(charpp_list[3])
                except:
                    await self.bot.say("The PP updating scripts are currently running. Please try again at a later time.")
                    return
            except:
                pass
        

        
        message = "These are your performance scores for your characters in HH20 P4:\n"
        message+= "Note: Feedback may be slightly inaccurate for characters with a low number of attempts.\n"
        for each in pp_list:
            message+= "**Name: {d[0]}**\n\t Deaths: {d[2]}\n\t Attempts: {d[3]}\n\t".format(d=each)
            if runs<50:
                await self.bot.send_message(author,"You do not have many runs. Please attend practice runs for some experience.")
            elif float(each[3])==0:
                message+= "This character has no runs yet.\n"
            elif float(each[1])>=7.5:
                message+= "This character is performing wonderfully! Keep up the good work!\n"
            elif float(each[1]) >=5:
                message+= "This character is doing decently, but you can still improve your death rate or do more runs.\n"
            elif float(each[1]) >=2.5:
                message+= "This character has a bit of deaths. Keep working on staying alive!\n"
            else:
                message+= "You have very few runs or are dying a bit too much. Speak to Rxkt for assistance so he can help you!\n"

        await self.bot.send_message(author,message)
        
    

    #Assigns the discord IDs for each given discord tag on the spreadsheet
    @commands.command(pass_context=True)
    async def assign_ids(self, ctx):
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
        
        

    #Notifies the nth raid on the spreadsheet (in case for raids created at the last moment.)
    @commands.command(pass_context=True)
    async def notify_raid(self, ctx,raid_index:int,*string):
        "Notifies the nth raid with a given string."

        server=ctx.message.server
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        wsh = sh.worksheet("InfoParse")
        wsh_raids = sh.worksheet("raid lineup")

        await self.bot.say("opened sheets . . .")
        raid_col = (raid_index-1)*5+1
        raid_row = 3 #first row is date, second is table header, 3rd is where characters begin.

        string_header= "This is a message for the raid on %s.\n" % wsh_raids.cell(1,raid_col).value

        characters = []
        for col in range(0,4):
            for row in range(0,10):
                name=wsh_raids.cell(row+raid_row,col+raid_col).value
                if name != '':
                    try:
                        cell = wsh.find(name)
                        id_target = wsh.cell(cell.row,5).value
                        member = server.get_member(id_target)
                        await self.bot.send_message( member, string_header + ' '.join(list(string)) )
                        #await self.bot.say("Sent message to %s, who is supposed to be on %s." % (member.name,name))
                    except:
                        await self.bot.say("Could not find %s on roster." % name)

                    
        await self.bot.say("Finished")
        
        

    
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

