import discord
import datetime
import time
import locale
import asyncio
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from discord.ext import commands
import webbrowser
import re
from __main__ import settings

scope = ['https://spreadsheets.google.com/feeds']

update_message = """
Hello, please update your answers to Amaterasu's 20 man form.

Form update rules: {}
""".format("https://docs.google.com/document/d/1Ws-GZwjvR2rOaZHfQue6qGipaH7d0ME3Jlvn-F1VJSA/edit")

no_member_found = """Unable to find you in the spreadsheet. Please go to

{}

to sign up for raids.""".format("https://docs.google.com/forms/d/e/1FAIpQLSfvhDhF1ko_iuCtTery4yjen2pzENy_vn5FpFOXmP8QZCF0iw/viewform")

class amabot(discord.Client):
    """Cog that scrapes Amaterasu's Raid spreadsheets!"""
    
    
    def __init__(self, bot):
        self.bot = bot

    def is_me(m):
        return m.author == client.user


    #When called, notifies all users who have not updated their spreadsheet within the past week to do so.
    @commands.command(no_pm=True,pass_context=True)
    async def notify(self, ctx,memberIndex = 2):
        """Mentions all people on the spreadsheet whose timestamps on their form responses are not updated for a period of time."""

        author = ctx.message.author
        roles = [x.name for x in author.roles]
        if "Power" not in roles:
            await self.bot.say("You don't have power to do that.")
            return

        server = ctx.message.server
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        wsh = sh.worksheet("InfoParse")
        await self.bot.say("opened sheet . . .")
        todayDate = datetime.datetime.today()
        
        cells=[x for x in wsh.get_all_values()[memberIndex-1:] if x[0]!=""]
        
        for each in cells:
            link = each[3]
            prevDate = datetime.datetime.strptime(each[0],"%m/%d/%Y %H:%M:%S")
            
            if True:#abs((prevDate-todayDate).days) > 4:
                member = server.get_member_named( each[1].strip(' ') )
                if member == None:
                    #given discord tag failed, let's try using the ID from assign_ids 
                    member = server.get_member( each[4].strip(' ') )

                if member != None:
                    if each[3] not in {'','undefined'}:
                        try:
                            await self.bot.send_message(member, "{}\n\n Your link is: {}\nIf you happen to have the wrong link, PM an officer on discord.".format(update_message,link))
                            #await self.bot.say("Messaged {}, line {}".format(each[2],memberIndex) )
                        except:
                            await self.bot.say("Could not message {}, line {}".format(each[2],memberIndex) )
                            
                    else:
                        try:
                            if link == '':
                                await self.bot.send_message(member, "You do not have a link now. PM Kyang.")
                                #await self.bot.say("doesnt have link")
                            else:
                                #await self.bot.say("undefined link")
                                await self.bot.send_message(member, "PM Kyang and tell him that you have an undefined link.")
                        except:
                            await self.bot.say("Could not message {}, line {}".format(each[2],memberIndex) )
                else:
                    await self.bot.say("{} not on roster.".format(each[2]) )
            memberIndex+=1


    
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


            
    #Sends a discord message to the user that called this function
    #with their respective performance ratings
    @commands.command(pass_context=True)
    async def pp(self, ctx):
        "Returns your performance rating of all your characters for raids in Phase 4."

        author = ctx.message.author

        #i think google's servers are wonky recently, let's catch this so my eyes don't bleed looking at command prompt
        try:
            credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
            gc = gspread.authorize(credentials)
            sh = gc.open("20man raid sheet")
        except:
            await self.bot.say("Unable to connect to Google Spreadsheets, please try again at a later time.")
            return
        
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
                message+= "This character has a bit of deaths or needs more runs. Try to get some more practice in!\n"
            else:
                message+= "You have very few runs or are dying a bit too much. Speak to Rxkt for assistance so he can help you!\n" 

        
        
        
        
        await self.bot.send_message(author,message)
        
        

    @commands.command(pass_context=True)
    async def points(self, ctx, *string):
        "Tells you your brooch points that you've earned."

        
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        wsh = sh.worksheet("InfoParse")
        boogaloo = gc.open("Harrowhold 2: Electric Boogaloo")
        points_wsh = boogaloo.worksheet("Points")

        string = ' '.join(string)
        if len(string)==0:
            #string= author
            author = ctx.message.author
            id_values=wsh.col_values(5)
            name_values=wsh.col_values(3)
            try:
                name = name_values[id_values.index(author.id)]
                points_values = points_wsh.col_values(2)
                main_char_values = points_wsh.col_values(3)
                await self.bot.say("Hi {}, you have {} brooch points.".format(name,points_values[main_char_values.index(name)]))
            except:
                await self.bot.say("You're not on the spreadsheet.")
        else:
            data= wsh.get_all_values()
            for row in data:
                for index in range(len(row)):
                    if re.match(row[index],string,re.IGNORECASE) and len(row[index]) == len(string):
                        name = row[2]
                        points_values = points_wsh.col_values(2)
                        main_char_values = points_wsh.col_values(3)
                        await self.bot.say("{} has {} brooch points.".format(name,points_values[main_char_values.index(name)]))

        
    #Assigns the discord IDs for each given discord tag on the spreadsheet
    @commands.command(no_pm=True,pass_context=True)
    async def assign_ids(self, ctx, memberIndex=2):
        "Updates the ID of every discord member on the spreadsheet."

        server=ctx.message.server
        
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        wsh = sh.worksheet("InfoParse")
        raw_data_wsh = sh.worksheet("RawData")
        await self.bot.say("opened sheet . . .")

        time_values = [x for x in wsh.col_values(1) if x!= ""]
        tags_values = wsh.col_values(2)
        id_values = wsh.col_values(5)
        tags_range = wsh.range(2,2,1+len(time_values),2)
        ids_range = raw_data_wsh.range(2,143,1+len(time_values),143)
        
        for index in range(len(time_values)):
            if tags_range[index].value != "":
                try:
                    member = server.get_member_named( tags_range[index].value.strip(' ') )
                    ids_range[index].value = member.id
                    print("Updated ID for {}".format(tags_range[index].value))
                except:
                    await self.bot.say("Could not get ID for {}".format(tags_range[index].value))
        
        
        wsh.update_cells(ids_range)
        

    #Notifies the nth raid on the spreadsheet (in case for raids created at the last moment.)
    @commands.command(no_pm=True,pass_context=True)
    async def notify_raid(self, ctx,raid_index:int,*string):
        "Notifies the nth raid with a given string."

        author = ctx.message.author
        roles = [x.name for x in author.roles]
        if "Power" not in roles:
            await self.bot.say("You don't have power to do that.")
            return

        server=ctx.message.server
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        wsh = sh.worksheet("InfoParse")
        wsh_raids = sh.worksheet("raid lineup")
        await self.bot.say("opened sheets . . .")
        raid_col = (raid_index-1)*5+1
        raid_row = 3 

        string_header= "\nThis is a message for the raid on {}.\n".format(wsh_raids.cell(1,raid_col).value)

        #get the characters we want to message
        characters = [x.value for x in wsh_raids.range(raid_row,raid_col,raid_row+9,raid_col+3) if x.value!= "" ]

        #download all the data we search through offline so we don't need to keep looking for it... in 2d list form
        all_values = wsh.get_all_values()
        data = [[i.lower() for i in x] for x in all_values[1:] if x[0] != ""]
        for name in characters:
            for row in data:
                if name.lower() in row:
                    try:
                        member = server.get_member(row[4])
                        await self.bot.send_message( server.get_member(row[4]), string_header + ' '.join(list(string)) ) 
                    except:
                        await self.bot.say("Unable to find {} on roster.".format(name) )

        await self.bot.say("Finished") 


    @commands.command(no_pm=True,pass_context=True)
    async def alert_reds(self, ctx):
        "Alerts all people with red cards to pay their fines if they haven't already."

        author = ctx.message.author
        roles = [x.name for x in author.roles]
        if "Power" not in roles:
            await self.bot.say("You don't have power to do that.")
            return

        server=ctx.message.server
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        boogaloo = gc.open("Harrowhold 2: Electric Boogaloo")
        wsh = boogaloo.worksheet("Infractions")

        await self.bot.say("going to open spreadsheet!")

        #let's load all of the data locally because calling cells takes WAY TOO LONG
        names = [name for name in wsh.col_values(11)[1:] if name!= '']
        red_cards = [int(card) for card in wsh.col_values(14)[1:2+len(names)]]
        transactions = [int(transaction) for transaction in wsh.col_values(15)[1:2+len(names)]]
        await self.bot.say("Opening sheet...")
        criminals = [names[index] for index in range(len(names)) if red_cards[index] > transactions[index] ]
        #await self.bot.say(criminals)

        #opening sheet with contact info
        sh = gc.open("20man raid sheet")
        wsh = sh.worksheet("InfoParse")
        for name in criminals:
            try:
                cell = wsh.find(name)
                id_target = wsh.cell(cell.row,5).value
                member = server.get_member(id_target)
                index = names.index(name)

                await self.bot.send_message( member,"You have {} red cards and have paid {} times.".format(red_cards[index], transactions[index] ) )
                await self.bot.send_message( member,"You currently owe guild bank {}k.".format( 75*(red_cards[index]-transactions[index]) ) )
                await self.bot.send_message( member, "Please pay ASAP or else you won't be able to participate in raids. Thanks!" )
            except:
                await self.bot.say("Unable to find {} on roster.".format(name) )
        await self.bot.say("Finished sending out notifications.")



    @commands.command(no_pm=True,pass_context=True)
    async def updatepp(self,ctx):
        "Updates pp because google scripts run really slowly..."

        author = ctx.message.author
        roles = [x.name for x in author.roles]
        if "Power" not in roles and "Spreadsheet" not in roles:
            await self.bot.say("You don't have power to do that.")
            return
        
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        #open ppsheet
        pp_wsh = sh.worksheet("PP Calculation2")

        
        await self.bot.say("opening pp sheets . . . ")
        
        #drag column P to column B, col Q to col C.
        col_P = [x for x in pp_wsh.col_values(16)[1:] if x != "" ]
        col_Q = pp_wsh.col_values(17)[1:1+len(col_P)]
        #print ( len(col_P) )

        col_B = pp_wsh.range(2,2,1+len(col_P),2)
        col_C = pp_wsh.range(2,3,1+len(col_P),3)
        
        clear_range = pp_wsh.range(2,2,pp_wsh.row_count,7)
        for cell in clear_range:
            cell.value=""
        pp_wsh.update_cells(clear_range)
        await self.bot.say("cleared all previous data")
        
        for index in range(len(col_P)):
            col_B[index].value = col_P[index]
            col_C[index].value = col_Q[index]
        pp_wsh.update_cells(col_B)
        pp_wsh.update_cells(col_C)

        await self.bot.say( "updated names & 30man pp")

        await self.bot.say("opening death sheets . . .")

        death_wsh = sh.worksheet("Death record2")
        
        #download names w/ attempts and deaths, strip with list comprehension get_all_values().
        #keep in mind this is 0based indexing in a list, unlike spreadsheet
        await self.bot.say("downloading cells...")
        cells = death_wsh.get_all_values()
        
        await self.bot.say("removing useless columns...")
        #a 2D list in the same format as the spreadsheet, but trimmed the top and empty columns. format: attempts,name,deaths
        raids = [[cells[j][i] for i in range(len(cells[j])) if i%6 in [2,4,5] ] for j in range(3,len(cells)) ] 

        #calculate pp
        await self.bot.say("creating table to upload...")
        #let's create the 2D list to project/upload back to google spreadsheets
        pp_chart = [[0 if i >1 else col_Q[j] if i>0 else col_P[j] for i in range(6)] for j in range(len(col_P)) ]

        await self.bot.say("transferring 30man PP to 20man PP...")
        for x in range(len(pp_chart)):
            #https://stackoverflow.com/questions/354038/how-do-i-check-if-a-string-is-a-number-float/23639915#23639915
            #bless stack overflow, to be honest
            if pp_chart[x][1].replace('.','',1).isdigit():
                #30man to 20man pp scaling as directed by kyang
                pp_chart[x][2]=(float(pp_chart[x][1])+10)/3
            else:
                pp_chart[x][2]=4
            

        await self.bot.say("calculating raid data and PP values...")


        
        #each raid is a column, not a row. our loop will be inverted...
        for x in range(int(len(raids[0] )/3)):
            current_raid = [row[x*3:(x+1)*3] for row in raids if row[x*3+1]!= '']
            #if len(current_raid)>0:
                #print (current_raid[0])

            #translating code from kyang's calculatepp here
            for each in current_raid:

                #incase they have nothing in the spreadsheet
                if each[2]=="":
                    each[2]=0
                if each[0]=="":
                    each[0]=0
                deaths= int(each[2])
                attempts = int(each[0])
                name = each[1]
                
                #we found the row for that person
                for item in pp_chart:
                    if re.match(name,item[0],re.IGNORECASE) and len(name)== len(item[0]):
                        item[2] = float(item[2])
                        
                        item[3]+= deaths
                        item[4]+= attempts
                        a = -0.5
                        
                        for j in range(1,attempts+1):
                            a+=deaths/attempts
                            if a>0:
                                item[2] *= 0.88
                                a-=1
                            item[2] = item[2]*.99 + .102
                            
        await self.bot.say("finalizing pp numbers and polishing the data...")
        #end scaling as directed by kyang, let's also polish up the board
        for item in pp_chart:
            item[2] = format(item[2]*2-5,'.2f')
            if item[3] ==0:
                item[5]="No death"
            else:
                item[5]=format(item[4]/item[3],'.2f')
        
            if item[4] ==0:
                item[5]="No attempt"

        update_range = pp_wsh.range(2,2,1+len(pp_chart),7)
        cols=len(pp_chart[0])
        for index in range(len(update_range)):
            #print(index)
            update_range[index].value=pp_chart[int(index/cols)][index%cols]
            
        await self.bot.say("uploading data")
        pp_wsh.update_cells(update_range)
        await self.bot.say("finished")


        
        """
code written by kyang from google script editor:

function updatepp(name, initialpp){
    var pp
    if (initialpp =="" || initialpp == "not enough runs")
        {pp = 4}
    else
        {pp = (initialpp+10)/3}
    var pplist = [pp, 0, 0]
    var numberofraids = countraids()
    for (i=1; i < numberofraids+1; i++){
        pplist = calculatepp(name, getdeathrange(i), pplist, i)
    }
    return pplist
}

function calculatepp(name, range, pplist, raidnumber){
    for (var i=1 ; i<range.getNumRows()+1;i++){
       if (range.getCell(i, 3).getValue().toString().toLowerCase()==name.toString().toLowerCase())
       {
         var ndeaths = range.getCell(i,4).getValue()
         var nattempts = range.getCell(i,1).getValue()
         pplist[1] = pplist[1]+ndeaths
         pplist[2] = pplist[2]+nattempts
         var formerpp = pplist[0]
         var a = -0.5
         for ( var j=1 ; j<range.getCell(i,1).getValue()+1; j++)
         {
          var a = a+ndeaths/nattempts
          if (a>0)
          {
            pplist[0] = pplist[0]-0.12*(pplist[0])
            a=a-1
          }
          pplist[0] = 0.01*(10-pplist[0])+0.002+pplist[0]   
         }
       }
  }
    return pplist
  }


        """


    @commands.command(no_pm=True,pass_context=True)
    async def addclear(self,ctx,name:str ,freebrooch:bool=False):
        "populates the boogaloo sheet's clear tab with the most recent raid that cleared\n use _addclear <name> true if they got a free brooch"

        author = ctx.message.author
        roles = [x.name for x in author.roles]
        if "Power" not in roles:
            await self.bot.say("You don't have power to do that.")
            return
        
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)

        await self.bot.say("opened death sheets . . .")
        sh = gc.open("20man raid sheet")
        death_wsh = sh.worksheet("Death record2")
        await self.bot.say("downloading cells...")
        cells = death_wsh.get_all_values()
        
        await self.bot.say("removing useless columns...")
        #a 2D list in the same format as the spreadsheet, but trimmed the top and empty columns. format: attempts,token,name,deaths
        raids = [[cells[j][i] for i in range(len(cells[j])) if i%6 in [2,3,4,5] ] for j in range(3,len(cells)) ]
        while raids[0][-4] == "":
            raids = [x[:-4] for x in raids]
        last_raid = [x[-4:] for x in raids]

        
        names = [x[2] for x in last_raid if x[1] != "" and int(x[1]) == 1]

        ##remove name if regex exists and put it at index 1. if not found, return and say char is not in there.
        matches= []
        for x in names:
            if (re.match(x,name,re.IGNORECASE) and len(name)== len(x) ):
                matches.append(x)
        if len(matches)==0:
            await self.bot.say("This person isn't in the raid.")
            return

        names = [x for x in names if not (re.match(x,name,re.IGNORECASE) and len(name)== len(x) ) ]
        names.insert(0,name)


        boogaloo = gc.open("Harrowhold 2: Electric Boogaloo")
        clear_wsh = boogaloo.worksheet("Raid Clears")
        row=clear_wsh.row_values(2)
        row = [x for x in row if x!=""]
        raid_index = len(row)+1
        update_range = clear_wsh.range(2,raid_index,25,raid_index)
        for index in range( len(names) ):
            update_range[index].value = names[index]
        if freebrooch:
            update_range[-1].value = name
        clear_wsh.update_cells(update_range)

        await self.bot.say("finished")
        
    @commands.command(no_pm=True,pass_context=True)
    async def addbonus(self,ctx):
        "gives bonus points to people who performed well"

        author = ctx.message.author
        roles = [x.name for x in author.roles]
        if "Power" not in roles:
            await self.bot.say("You don't have power to do that.")
            return

        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)

        sh = gc.open("20man raid sheet")
        death_wsh = sh.worksheet("Death record2")
        cells = death_wsh.get_all_values()
        raids = [[cells[j][i] for i in range(len(cells[j])) if i%6 in [2,3,4,5] ] for j in range(3,len(cells)) ]
        while raids[0][-4] == "":
            raids = [x[:-4] for x in raids]
        last_raid = [x[-4:] for x in raids]

        #give points to those who have the max # of attempts and have never died
        attempts = max([int(x[0]) for x in last_raid if x[0]!=''])
        zerodeathclub = []
        for x in last_raid:
            if x[0] != '' and int(x[0])==attempts and int(x[1])==1 and int(x[3])==0:
                zerodeathclub.append(x[2])

        info_wsh = sh.worksheet("InfoParse")
        info = info_wsh.get_all_values()
        print (zerodeathclub)
        #replace the names of the alts with their main
        for index in range(len(zerodeathclub)):
            name=zerodeathclub[index]
            for row in info:
                for each in row:
                    if re.match(name,each,re.IGNORECASE) and len(name)== len(each):
                        zerodeathclub[index]=row[2]

        #open the points sheet and give out the bonus points
        boogaloo = gc.open("Harrowhold 2: Electric Boogaloo")
        points_wsh = boogaloo.worksheet("Points")
        names = [x for x in points_wsh.col_values(3) if x != ""][1:]
        update_range = points_wsh.range(3,6,len(names)+2,6)

        if len(update_range) != len (names):
            await self.bot.say("please make sure there are enough 0's on the brooch points sheet")
            return
        for name in zerodeathclub:
            for index in range( len(update_range) ):
                if re.match(name,names[index], re.IGNORECASE) and len(name) == len(names[index]) :
                    update_range[index].value = float(update_range[index].value)+ 0.25
                    await self.bot.say("gave 0.25 to "+name)
        points_wsh.update_cells(update_range)
                    
        await self.bot.say("finished")


    @commands.command(no_pm=True,pass_context=True)
    async def setupraid(self,ctx,raid_index:int):
        "Sets up the raid in the death record spreadsheet."

        author = ctx.message.author
        roles = [x.name for x in author.roles]
        if "Power" not in roles:
            await self.bot.say("You don't have power to do that.")
            return
        
        server=ctx.message.server
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        death_wsh = sh.worksheet("Death record2")
        wsh_raids = sh.worksheet("raid lineup")
        await self.bot.say("opened sheets . . .")
        raid_col = (raid_index-1)*5+1
        raid_row = 3 

        
        #get the characters we want
        characters = [x.value for x in wsh_raids.range(raid_row,raid_col,raid_row+9,raid_col+3) if x.value not in ["","Raid Type","Phase 1-4","Phase 4 Practice"]]
        #get the column number
        rowoneary = death_wsh.row_values(1)
        rowone = [rowoneary[i] for i in range(len( rowoneary )) if i%6==2]
        col_num= rowone.index("#DIV/0!")*6+3

        #update attempts
        attempt_range = death_wsh.range(4,col_num,23,col_num)
        for each in attempt_range:
            each.value = 0
        death_wsh.update_cells(attempt_range)
        #update names
        name_range=death_wsh.range(4,col_num+2,23,col_num+2)
        if len(name_range)!=20:
            await self.bot.say("Not all 20 spots are filled.")
        
        for index in range(len(name_range)):
            name_range[index].value = characters[index]
        death_wsh.update_cells(name_range)
        #update deaths
        death_range = death_wsh.range(4,col_num+3,23,col_num+3)
        for each in death_range:
            each.value =0
        death_wsh.update_cells(death_range)
            
        await self.bot.say("Finished")


    @commands.command(no_pm=True,pass_context=True)
    async def addattempt(self,ctx,*string):
        "Adds an attempt to most recent raid in death record spreadsheet."
        
        author = ctx.message.author
        roles = [x.name for x in author.roles]
        if "Power" not in roles:
            await self.bot.say("You don't have power to do that.")
            return
        
        server=ctx.message.server
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        death_wsh = sh.worksheet("Death record2")

        #get the column number
        rowary = death_wsh.row_values(4)
        rowfour = [rowary[i] for i in range(len( rowary )) if i%6==2 and rowary[i]!=""]
        col_num=len(rowfour)*6-3

        toggle_values = [x for x in death_wsh.col_values(col_num+1)[3:] if x!=""]
        
        attempt_range = death_wsh.range(4,col_num,3+len(toggle_values),col_num)
        for index in range(len(toggle_values)):#each in attempt_range:
            if int(toggle_values[index])!=0:
                if attempt_range[index].value!="":
                    attempt_range[index].value=int(attempt_range[index].value)+1
                else:
                    attempt_range[index].value=1
        death_wsh.update_cells(attempt_range)
        deaths = " ".join(string)
        
        if len(deaths)>0:
            death_range = death_wsh.range(4,col_num+2,3+len(toggle_values),col_num+3)
            death_list = [x.strip(' ') for x in deaths.split(',')]
            for dead_person in death_list:
                for index in range(len(death_range)):
                    name = death_range[index].value
                    #must convert to string since cells may hold ints, not strings
                    if re.match(str(dead_person),str(name),re.IGNORECASE) and len(str(name))== len(str(dead_person)):
                        death_range[index+1].value = int(death_range[index+1].value)+1
                        
            death_wsh.update_cells(death_range)

        await self.bot.say("finished")

    @commands.command(no_pm=True,pass_context=True)
    async def sub(self,ctx,*string):
        "Subs 1 person out for another in the most recent raid in the death record spreadsheet."

        author = ctx.message.author
        roles = [x.name for x in author.roles]
        if "Power" not in roles:
            await self.bot.say("You don't have power to do that.")
            return

        string = ' '.join(string)
        names = [x.strip(' ') for x in string.split(',')]
        if len(string)==0 or ',' not in string:
            await self.bot.say("Please enter subbed,subber")
            return
        
        server=ctx.message.server
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        death_wsh = sh.worksheet("Death record2")

        #get the column number
        rowary = death_wsh.row_values(4)
        rowfour = [rowary[i] for i in range(len( rowary )) if i%6==2 and rowary[i]!=""]
        col_num=len(rowfour)*6-3

        toggle_values = [x for x in death_wsh.col_values(col_num+1)[3:] if x!= ""]
        #range will have 1 extra row to "prepare" for new person
        toggle_range = death_wsh.range(4,col_num+1,4+len(toggle_values),col_num+1)
        name_range = death_wsh.range(4,col_num+2,4+len(toggle_values),col_num+2)
        death_range = death_wsh.range(4,col_num+3,4+len(toggle_values),col_num+3)

        #check subbed
        #since regex can't match str length well
        name_values = [x.value for x in name_range if len(x.value) == len(names[0])]
        r = re.compile(names[0],re.IGNORECASE)
        newlist = list(filter(r.match, name_values))
        if len(newlist)==0:
            await self.bot.say("{} not found.".format(names[0]))
            return
        subbed=newlist[0]
        #check subber
        name_values = [x.value for x in name_range if len(x.value) == len(names[1])]
        r = re.compile(names[1],re.IGNORECASE)
        newlist = list(filter(r.match, name_values))
        
        name_values = [x.value for x in name_range]
        if len(newlist)==0:
            #add
            toggle_range[name_values.index(subbed)].value = 0
            toggle_range[-1].value=1
            name_range[-1].value=names[1]
            death_range[-1].value=0
            death_wsh.update_cells(toggle_range)
            death_wsh.update_cells(death_range)
            death_wsh.update_cells(name_range)
        else:
            toggle_range[name_values.index(subbed)].value = 0
            toggle_range[name_values.index(newlist[0])].value = 1
            death_wsh.update_cells(toggle_range)
        await self.bot.say("finished")

    @commands.command(no_pm=True,pass_context=True)
    async def status(self,ctx):
        "Shows death record of current raid."
        
        author = ctx.message.author
        roles = [x.name for x in author.roles]
        if "Power" not in roles and "Spreadsheet" not in roles:
            await self.bot.say("You don't have power to do that.")
            return
        
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        death_wsh = sh.worksheet("Death record2")
        #get the column number
        rowary = death_wsh.row_values(4)
        rowfour = [rowary[i] for i in range(len( rowary )) if i%6==2 and rowary[i]!=""]
        col_num=len(rowfour)*6-3

        
        num_players = len([x for x in death_wsh.col_values(col_num+1)[3:] if x!= ""])
        #await self.bot.say(num_players)
        raid_range=death_wsh.range(3,col_num,3+num_players,col_num+3)
        output="Current raid status:\n```"
        longest_name_len = max([len(x.value) for x in death_wsh.range(4,col_num+2,3+num_players,col_num+2)])
        for index in range(len(raid_range)):
            output+="{message:{fill}{align}{width}}".format(message=raid_range[index].value,fill=' ',align='^' if index%4==0 else '<',width=[9,3,longest_name_len+1,1][index%4] )
            if index%4==3:
                output+="\n"
        output+="```"
        await self.bot.say(output)

    ##-----form management functions-----
    
    def _open_data(self):
        "Opens spreadsheet"
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        wsh = sh.worksheet("RawData(BETA)")
        return wsh

    async def _log_msg(self, string):
        await self.bot.send_message(discord.Object("375455571526156288"),string)
    
    def _update_time(self,wsh,row):
        wsh.update_cell(row,1,time.strftime("%m/%d/%Y %H:%M:%S"))

    
    @commands.group(pass_context=True)
    async def form(self,ctx):
        "Function that will replace & overhaul the form-spreadsheet system"
        if ctx.invoked_subcommand is None:
            await self.bot.send_cmd_help(ctx)
            return

    @form.command(pass_context=True,name="status")
    async def _form_status(self,ctx):
        "Shows you your current raid signup status."

        wsh = self._open_data()
        author= ctx.message.author
        ids = [x for x in wsh.col_values(2) if x!= ""]
        if author.id not in ids:
            await self.bot.say("You are not on the spreadsheet.")
        else:
            row_num = ids.index(author.id)+1
            row = wsh.row_values(row_num)
            string = "*Last updated: {}*```\n".format(row[0])
            string += "Characters:\n"
            string += "Main: {},{}\n".format(row[2],row[3])
            for index in range(16):
                if row[index*2+4] != "":
                    string+= "Alt {}: {},{}\n".format(index+1,row[index*2+4],row[5+2*index])
            string += "\nHours available:\n"
            for index in range(7):
                string+= "{}: {}\n".format(["Tues","Wed ","Thur","Fri ","Sat ","Sun ","Mon "][index],row[36+index])
            string +="```"
            await self.bot.say(string)
            
    
    @form.command(pass_context=True,name="add")
    async def _form_add(self,ctx,name:str,char_class:str):
        "Add a character."

        if char_class.lower() not in ["brawler","lancer","warrior","slayer","archer","reaper","berserker","sorcerer","valkyrie","gunner","ninja","priest","mystic"]:
            await self.bot.say("Please enter a valid class.")
            return
        
        #update time
        
        wsh = self._open_data()
        author = ctx.message.author
        ids = [x for x in wsh.col_values(2) if x!= ""]
        if author.id not in ids:
            #set this as their main
            row_num = len(ids)+1
            self._update_time(wsh,row_num)
            update_range = wsh.range(row_num,1,row_num,4)
            update_range[2].value = name
            update_range[1].value = author.id
            update_range[3].value = char_class.lower().title()
            wsh.update_cells(update_range)
            await self.bot.say("Added {}.".format(name))
            await self._log_msg("{}: :pencil: **SIGN UP:** {}({})".format(author.name,name,char_class))
        else:
            row_num = ids.index(author.id)+1
            self._update_time(wsh,row_num)
            update_range = wsh.range(row_num,5,row_num,36 )
            index = len([x.value for x in update_range if x.value != ""])
            if index >= len(update_range):
                await self.bot.say("You cannot add more alts :(")
                return
            #check if in graveyard.
            credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
            gc = gspread.authorize(credentials)
            boogaloo = gc.open("Harrowhold 2: Electric Boogaloo")
            roster_wsh = boogaloo.worksheet("Roster")
            #flatten 2d array to list
            names = [i.lower() for x in roster_wsh.get_all_values() for i in x]
            if name.lower() in names:
                await self.bot.say("That character is already in the roster.")
                return

            #####################
            update_range[index].value=name
            update_range[index+1].value=char_class.lower().title()
            wsh.update_cells(update_range)
            await self.bot.say("Added {}.".format(name))
            await self._log_msg("{}: :o: **ADD:** {}({})".format(author.name,name,char_class))
            return
            
        
    @form.command(pass_context=True,name="remove")
    async def _form_remove(self,ctx,name:str):
        "Remove a character."

        wsh = self._open_data()
        author = ctx.message.author
        ids = [x for x in wsh.col_values(2) if x!= ""]
        if author.id not in ids:
            await self.bot.say("You are not on the spreadsheet.")
        else:
            row_num = ids.index(author.id)+1
            #update time
            self._update_time(wsh,row_num)    
            update_range = wsh.range(row_num,1,row_num,len(wsh.row_values(row_num)) )
            
            alts_copy = []
            for each in update_range[4:36]:
                alts_copy.append(each.value)
                each.value=""
            names = [x.lower() for x in alts_copy]
            if name.lower() not in names:
                await self.bot.say("That character isn't one of your alts.")
                return
            index = [x.lower() for x in alts_copy].index(name.lower())
            alts_copy.pop(index)
            alts_copy.pop(index)
            for each in update_range[4:36]:
                if len(alts_copy)>0:
                    each.value=alts_copy.pop(0)
            wsh.update_cells(update_range)
            await self.bot.say("Removed {}. Remember to _form graveyard so you keep the brooch points.".format(name))
            await self._log_msg("{}: :x: **REMOVE:** {}".format(author.name,name))

    ##reminder that just renaming does not cut it. You must change in death record. Remind to add to graveyard.
    @form.command(pass_context=True,name="rename")
    async def _form_rename(self,ctx,old_name:str,new_name:str):
        "Rename a character."

        start_time = time.time()
        wsh = self._open_data()
        author = ctx.message.author
        ids = [x for x in wsh.col_values(2) if x!= ""]
        if author.id not in ids:
            await self.bot.say("You are not on the spreadsheet.")
        else:
            row_num = ids.index(author.id)+1
            #update time
            self._update_time(wsh,row_num)
            #check if they want to rename main, then alts
            update_range = wsh.range(row_num,3,row_num,36 )
            names = [update_range[index].value.lower() for index in range(len(update_range)) if index%2==0 and update_range[index].value!=""]
            if old_name.lower() not in names:
                await self.bot.say("Cannot find {}.".format(old_name))
                return
            else:
                #rename in death record.
                credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
                gc = gspread.authorize(credentials)
                sh = gc.open("20man raid sheet")
                death_wsh = sh.worksheet("Death record2")
                _useless_row = death_wsh.row_values(4)
                name_cols = [x for x in range(len(_useless_row)) if x%6==4 and _useless_row[x]!=""]
                num_rows = len(death_wsh.col_values(1))

                #the col number of appearances of old_name
                appearances = []
                data = death_wsh.get_all_values()
                for i in range(len(data)):
                    for j in name_cols:
                        if data[i][j].lower()==old_name.lower():
                            appearances.append(j)
                
                appearances = [x+1 for x in appearances]
                to_update=[]
                for col_num in appearances:
                    #check which row the name falls in each column it appears in
                    raid_range = death_wsh.range(4,col_num,35,col_num) #35=len(death_wsh.col_values(0)), 1 less call
                    for each in raid_range:
                        if each.value.lower()==old_name.lower():
                            each.value=new_name
                            to_update.append(each)

                
                update_range[2*names.index(old_name.lower())].value=new_name
                death_wsh.update_cells(to_update)
                wsh.update_cells(update_range)
                await self.bot.say("Renamed {} to {}. Remember to _form graveyard so you keep the brooch points.".format(old_name,new_name))
                await self._log_msg("{}: :warning: **RENAME:** {}->{}".format(author.name,old_name,new_name))

                await self.bot.say("--- Execution time: %s seconds ---" % (time.time() - start_time))
                return

    @form.command(pass_context=True,name="graveyard")
    async def _form_graveyard(self,ctx,operation:str,name:str):
        "Add/removes a character to the graveyard. For retired/renamed/removed characters to retain the clear count."

        author = ctx.message.author
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        info_wsh = sh.worksheet("InfoParse")
        boogaloo = gc.open("Harrowhold 2: Electric Boogaloo")
        roster_wsh = boogaloo.worksheet("Roster")

        ids = info_wsh.col_values(5)
        if author.id not in ids:
            await self.bot.say("You are not on the spreadsheet.")
            return
        if operation.lower() not in ["add","remove"]:
            await self.bot.say("Improper operation. Please select add or remove.")
            return

        names = [i.lower() for x in roster_wsh.get_all_values() for i in x]
        main = info_wsh.cell(ids.index(author.id)+1,3).value
        mains = roster_wsh.col_values(1)
        row = mains.index(main)+1
        row_values = roster_wsh.row_values(row)
        #the graveyard is at the end of the row
        row_values.reverse()
        reverse_index = row_values.index("")

        if operation.lower()=="add":
            if name.lower() in names:
                await self.bot.say("That character is already in the roster. Please remove it before adding it to the graveyard.")
                return
            roster_wsh.update_cell(row,len(row_values)-reverse_index,name)
            await self.bot.say("Added {} to graveyard.".format(name))
            await self._log_msg( "{}: :skull: **GRAVEYARD->add**: {}".format(author.name,name))

        else:#operation=="remove"
            if name.lower() not in names:
                await self.bot.say("{} does not exist in your graveyard.".format(name))
                return
            temp_grave = []  
            update_range = roster_wsh.range(row,len(row_values)-reverse_index+1,row,len(row_values)-1 )
            for each in update_range:
                if each.value.lower() != name.lower():
                    temp_grave.append(each.value)
                each.value=""
            update_range.reverse()
            temp_grave.reverse()
            for index in range(len(update_range)):
                if len(temp_grave)>0:
                    update_range[index].value=temp_grave.pop(0)
            roster_wsh.update_cells(update_range)
            await self.bot.say("Removed {} from the graveyard.".format(name))
            await self._log_msg( "{}: <:blobhyperthink:361641325193330690> **GRAVEYARD->remove**: {}".format(author.name,name))
        
        
    @form.command(pass_context=True,name="times")
    async def _form_times(self,ctx,day:str,*free_times):
        "0=Noon, 12=Midnight, busy=Busy all day. You can use ranges (0-12) or commas (12,1,2,3,...)"

        days = ["tuesday","wednesday","thursday","friday","saturday", "sunday","monday"]
        word_search = [x for x in days if day in x]
        if len(word_search) != 1:
            #either no search found, or not a unique string, i.e "day" or "ay"
            await self.bot.say("Please enter a correct day of the week.")
            return
        index = days.index(word_search[0])
        #remove all those useless spaces
        free_times = " ".join(free_times)
        free_times = free_times.replace(" ","")
        hour_ranges = free_times.split(",")
        hours = []
        for each in hour_ranges:
            if "-" in each:
                _range = each.split("-")
                if "" in _range:
                    await self.bot.say("Please enter a valid range <:blobfacepalm:324810235447607300>")
                    return
                for number in range(int(_range[0]),int(_range[1])+1):
                    try:
                        #check if they gave us time values in non-ascending order
                        if len(hours) > 0 and hours[len(hours)-1] > number:
                            await self.bot.say("Please enter the hours in order.")
                            return
                        hours.append(number)
                    except:
                        await self.bot.say("Please enter hours in a correct format.")
                        return
            else:
                try:
                    #busy all day, no numbers in string
                    if free_times.lower()=="busy":
                        self._update_time(wsh,row_num)
                        wsh.update_cell(row_num,37+index,"Busy")
                        await self.bot.say("Updated to BUSY on {}.".format(word_search[0]))
                        return
                    #check if they gave us time values in non-ascending order
                    if len(hours) > 0 and hours[len(hours)-1] > int(each):
                        await self.bot.say("Please enter the hours in order.")
                        return
                    
                    hours.append(int(each))
                except:
                    await self.bot.say("Please enter hours in a correct format.")
                    return

        
        hours_string=", ".join([["Noon","1:00 PM","2:00 PM","3:00 PM","4:00 PM","5:00 PM","6:00 PM","7:00 PM","8:00 PM","9:00 PM","10:00 PM","11:00 PM","Midnight"][x] for x in hours])
        
        wsh = self._open_data()
        author = ctx.message.author
        ids = [x for x in wsh.col_values(2) if x!= ""]
        if author.id not in ids:
            await self.bot.say("You are not on the spreadsheet.")
        else:
            row_num = ids.index(author.id)+1
            #update time
            self._update_time(wsh,row_num)
            #await self.bot.say(hours_string)
            wsh.update_cell(row_num,37+index,hours_string)
            await self.bot.say("Updated times on {}.".format(word_search[0]))
    ##-----functions for my lazy butt-----

            


            
    @commands.command(pass_context=True)
    async def raidcommands(self,ctx):
        "returns a list of useful commands for raid"

        #the ugliest string ever
        
        string="```\n"
        string+="_pp: Returns your performance for your characters.\n"
        string+="_link: Makes grillbot DM you your link.\n"
        string+="_lineup: Returns a link to the raid lineup sheet.\n"
        string+="_points [character]: Returns the amount of brooch points you have.\n"
        string+="\n"
        author = ctx.message.author
        roles = [x.name for x in author.roles]
        if "Power" in roles:
            string+="<> is required parameter, [] is optional.\n"
            string+="_notify [index]: Notifies people to update their HH20 form.\n"
            string+="_notify_raid <index> <string>: Notifies people in the <index>th raid with <string>.\n"
            string+="_alert_reds: Wake-up call to people with red-cards.\n"
            string+="_updatepp: Updates PP in a blink of an eye. or at least better than google spreadsheets\n"
            string+="_addclear <brooch winner> [true if free brooch]: Updates last clear on Boogaloo sheet.\n"
            string+="_addbonus: Gives bonus brooch points to 0-death-club people.\n"
            string+="_setupraid <index>: Creates the raid in death record spreadsheet.\n"
            string+="_addattempt [name1,name2,name3...]: Adds 1 attempt to most recent raid. Deaths optional.\n"
            string+="_sub <subbed>,<subber>: Manages subbing on most recent raid on death record sheet.\n"
            string+="_status: Shows the current raid in death record sheet.\n"
        
        string+="```"
        await self.bot.say(string)

    
    #spits out a certain cell, test func
    async def spit_cell(self,ctx,row,col,*sh_name):
        "Spits out a certain cell on a certain spreadsheet"
        
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        wsh = sh.worksheet("InfoParse")

        await self.bot.say("opening {}".format( ' '.join(list(sh_name))) )
        wsh_ = sh.worksheet(' '.join(list(sh_name)) )
        
        await self.bot.say(wsh_.cell(row,col).value)
        
    
    # Lists all raids not yet occurred. useless for now, may want to put this into a PM.
    @commands.command()
    async def raids(self):
        "Lists all the raid that have not yet occurred."

        await self.bot.say("Searching the spreadsheets for raids...")
        
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        wsh = sh.worksheet("raid lineup")

        raidIndex = 1
        raidTimes="This week's raids:\n"

        raids = [x for x in wsh.row_values(1) if x not in ['','EDT']]
        raid_types = [x for x in wsh.row_values(21) if x != '']
        
        for index in range(len(raids)):
            raidTimes+="{}) {}\t *{}*\n".format(index+1,raids[index],raid_types[index])

        await self.bot.say(raidTimes)

    @commands.command(pass_context=True)
    async def myraids(self,ctx):
        "Returns your raids."
        locale.setlocale(locale.LC_TIME, "")
        
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        #await self.bot.say("Opening Amaterasu's raid sheets...")
        sh = gc.open("20man raid sheet")
        wsh = sh.worksheet("raid lineup")
        info_wsh = sh.worksheet("InfoParse")
        #await self.bot.say("Opening Manifest's raid sheets...")
        mani_sh = gc.open("HH20 Raids - Manifest")
        mani_wsh = mani_sh.worksheet("Raid Schedule")
        #await self.bot.say("Searching both sheets...")
        author = ctx.message.author
        row_num = info_wsh.col_values(5).index(author.id)+1
        row = info_wsh.row_values(row_num)
        chars = [row[x].lower() for x in range(len(row)) if (x==2 or x>4) and row[x]!=""]

        string="**RAIDS:**\n"
        #ama raid
        num_raids=len([x for x in wsh.row_values(2) if x!=""])/4
        
        values = wsh.get_all_values()
        raids = [[i[x].lower() for x in range(len(i)) if x%5!= 4] for i in values[2:12] ] #cut out 5th column,lowercase
        filtered_raids = []
        for each in raids:
            matches = [x for x in each if x in chars]
            for name in matches:
                index = each.index(name)
                #append [name,time,raid_type,ama=0/mani=1]
                filtered_raids.append( [name, values[0][int(index/4)*5] , values[8][int(index/4)*5 + 3] , 0] )

        
        
        #mani raid
        mani_raids="**MANIFEST RAIDS:**\n"
        num_raids=len([x for x in mani_wsh.row_values(3) if x!=""])/4
        values = mani_wsh.get_all_values()
        raids = [[i[x].lower() for x in range(len(i)) if x%5!= 4] for i in values[2:12] ] #cut out 5th column,lowercase
        for each in raids:
            matches = [x for x in each if x in chars]
            for name in matches:
                #we need to normalize their spreadsheet's time format with ours...
                index = each.index(name)
                date_day = values[0][int(index/4)*5].split(" ")
                day_value = ["Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"][["Tue","Wed","Thu","Fri","Sat","Sun","Mon"].index(date_day[1].strip(" "))]
                date_value = datetime.datetime.strptime( date_day[0].strip(" ") , "%m/%d").strftime("%b %d, ")
                hour = values[0][int(index/4)*5 +1].split("-")[0]
                if ":" not in hour:
                    hour = hour.replace("PM",":00PM")
                hour = hour.replace("PM","p")
                raid_type = values[0][int(index/4)*5 +3].replace("p","Phase ")
                filtered_raids.append([ name, (day_value+", " +date_value+hour).strip(" "), raid_type, 1])
        
        filtered_raids.sort(key=lambda x: datetime.datetime.strptime(x[1], "%A, %b %d, %I:%Mp"))

        next_raid_index = 0
        for each in filtered_raids:
            strikeout= ""
            if datetime.datetime.strptime(each[1]+"m","%A, %b %d, %I:%M%p").replace(year=2017) < datetime.datetime.now():
                strikeout="~~"
                next_raid_index+=1
                
            string+="{}:regional_indicator_{}:{}*({})*: **{}**{}\n".format(strikeout,"a" if each[3]==0 else "m",each[1],each[2],each[0].title(),strikeout)

        time_remaining = datetime.datetime.strptime(filtered_raids[next_raid_index][1]+"m","%A, %b %d, %I:%M%p").replace(year=2017) - datetime.datetime.now()
        secs = time_remaining.seconds % 60
        mins = int(time_remaining.seconds/60)%60
        hours = int(time_remaining.seconds/3600)
        days = time_remaining.days
        raid_type = ":regional_indicator_{}:{}".format("a" if filtered_raids[next_raid_index][3]==0 else "m","materasu" if filtered_raids[next_raid_index][3]==0 else "anifest")
        string+="*Next raid by {} is in {}{}{}{}{}{}{} seconds.*".format( raid_type,
                                                                        days if days>0 else ""," days, " if days>0 else "",
                                                                        hours if hours>0 else ""," hours, " if hours>0 else "",
                                                                        mins if mins>0 else ""," minutes, " if mins>0 else "", secs )
        
        await self.bot.say(string)
    @commands.command(pass_context=True)
    async def lineup(self,ctx):
        "returns boogaloo"

        if ctx.message.author.id == settings.owner:
            webbrowser.open('https://docs.google.com/spreadsheets/d/18_dQWsIQy35jxuf7daSxQGfRFRs3_HJrlPqtZ5Pm6iE/', new=2)
            await self.bot.say("Opened the spreadsheet Electric Boogaloo. <:blobuwu:351058556449062912>")
            return
        
        await self.bot.say("<https://docs.google.com/spreadsheets/d/18_dQWsIQy35jxuf7daSxQGfRFRs3_HJrlPqtZ5Pm6iE/>")

    @commands.command(pass_context=True)
    async def addcols(self,ctx,numcols):
        "Add cols"

        await self.bot.say("opening sheets")
        credentials = ServiceAccountCredentials.from_json_keyfile_name('data/spreadsheet/Ama30man-baa10dc23211.json', scope)
        gc = gspread.authorize(credentials)
        sh = gc.open("20man raid sheet")
        wsh = sh.worksheet("Death record2")
        wsh.add_cols(numcols)
        await self.bot.say("finished")
        
    @commands.command(pass_context=True)
    async def main_sheet(self,ctx):
        "returns pp sheet"

        if ctx.message.author.id == settings.owner:
            webbrowser.open('https://docs.google.com/spreadsheets/d/1gRFcj8uDgchpapSbimt_JyRRAbdGPY4lhmqM878cXl8/edit#gid=21311540', new=2)
            await self.bot.say("Opened the spreadsheet 20man raid sheet. <:blobuwu:351058556449062912>")
        else:
            await self.bot.say("You do not have permissions. <:blobsob:317541132252872704>")

    @commands.command(pass_context=True)
    async def mechanics(self,ctx):
        "returns mechanics list"

        if ctx.message.author.id == settings.owner:
            webbrowser.open('https://docs.google.com/spreadsheets/d/1kx-ODqPh4039d2ijzxefzJjjKcFJ8ovhQfTS2rQcD_I/edit#gid=1985626731',new=2)
            await self.bot.say("Opened the spreadsheet 20man mechs. <:blobuwu:351058556449062912>")
        else:
            await self.bot.say("You do not have permissions. <:blobsob:317541132252872704>")

def setup(bot):
    bot.add_cog(amabot(bot))

