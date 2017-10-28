import discord
import datetime
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
    @commands.command(pass_context=True)
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

        #probably dont need this anymore
        '''
        
        if len(pp_list)!= 0:
            await self.bot.send_message(author,message)
        else:
            await self.bot.send_message(author,"The system is currently under maintenance. Check back later.")
        '''

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
    @commands.command(pass_context=True)
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
    @commands.command(pass_context=True)
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
        data = [x for x in wsh.get_all_values()[1:] if x[0] != ""]
        for name in characters:
            r = re.compile(name,re.IGNORECASE)
            for each in data:
                results = filter(r.search, each)
                if len( list(results) ) > 0:
                    try:
                        member = server.get_member(each[4])
                        await self.bot.send_message( server.get_member(each[4]), string_header + ' '.join(list(string)) ) 
                    except:
                        await self.bot.say("Unable to find {} on roster.".format(name) )
            
        await self.bot.say("Finished") 


    @commands.command(pass_context=True)
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



    @commands.command(pass_context=True)
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


    @commands.command(pass_context=True)
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
        
    @commands.command(pass_context=True)
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


    @commands.command(pass_context=True)
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
        for index in range(len(name_range)):
            name_range[index].value = characters[index]
        death_wsh.update_cells(name_range)
        #update deaths
        death_range = death_wsh.range(4,col_num+3,23,col_num+3)
        for each in death_range:
            each.value =0
        death_wsh.update_cells(death_range)
            
        await self.bot.say("Finished")


    @commands.command(pass_context=True)
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

    @commands.command(pass_context=True)
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

    @commands.command(pass_context=True)
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
        raid_range=death_wsh.range(3,col_num,2+num_players,col_num+3)
        output="Current raid status:\n"
        for index in range(len(raid_range)):
            output+="{}\t".format(raid_range[index].value)
            if index%4==3:
                output+="\n"
        await self.bot.say(output)
    #######################################
    #functions for my lazy butt

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
    @commands.command(pass_context=True)
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
    async def lineup(self,ctx):
        "returns boogaloo"

        if ctx.message.author.id == settings.owner:
            webbrowser.open('https://docs.google.com/spreadsheets/d/18_dQWsIQy35jxuf7daSxQGfRFRs3_HJrlPqtZ5Pm6iE/', new=2)
            await self.bot.say("Opened the spreadsheet Electric Boogaloo. <:blobuwu:351058556449062912>")
            return
        
        await self.bot.say("<https://docs.google.com/spreadsheets/d/18_dQWsIQy35jxuf7daSxQGfRFRs3_HJrlPqtZ5Pm6iE/>")

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

