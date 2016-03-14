## Match History Analyzer  
****
Match History Analyzer is an Excel Spreadsheet that takes data from your games and calculates stats such as your CS/min and Gold/min on a per game basis, but also over all games you have played. It shows you these stats for all your champions, but you can also filter for certain champions, even by role.  
It also calculates your winrate and playrate vs certain matchups, and allows you to add notes to these matchups. 

****
#### What do you need
****
Microsoft Office Excel (2010 or newer) installed, on Windows. Mac is not supported, since it uses a different version of VBA than Windows. (sorry)

#### [Download link](https://github.com/ikhebgeenaccount/Excel-Match-History-Analyzer/releases/tag/1.3)

****
### Supporting and contributing
****
If you'd like to help with this project, contact me via reddit, [/u/ikhebgeenaccount](http://reddit.com/u/ikhebgeenaccount). If you want to show your appreciation for this project and help me out so I can keep spending time on it, [you can donate a few bucks](https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=ESHNNFQLF8RF8).

****
### How does it work
****
When you have downloaded the spreadsheet, and have Microsoft Excel installed on your PC, you can open the file. You will be prompted by a messagebox telling you macro's have been disabled. These have to be enabled to make this work, so click *enable*. After that, you need to fill in some information. In cells C3 and C5 fill in respectively your summoner name and region. After that, you want to get an API key from http://developer.riotgames.com and fill it in cell C6.  
Then you're good to go. When you hit update, your rank will be filled in, your top 6 champions and your last 10 games. Stats will be calculated and you can see those in the table on the "Front". You can also fill in cells G20 - G24 with champion names from which you want to get individual stats. The same goes for cells D18 and E18.

****
### Matchups
****  
To see your performance against certain matchups go to the "Matchups" page and fill in the champion about which you want to see it in cell B1. Mind that this data will not be very usefull until you have some more matches in your spreadsheet than the first 10.

****
### Then what
****
After you have updated the spreadsheet for the first time, you can update it after each game you play (or after multiple games) and it will insert those new games. Old games will not be deleted. It will take some time to get useful data from your spreadsheet, since you need more than a few games to be able to draw conclusions. After about 20 games you should be able to see some patterns in your stats, and be able to determine on what fields you're lacking.

****
### Migrating games from one version to another
****
First of all, nice you downloaded it :) Now, you probably don't want to start over with collecting your games. Dont worry tho, you won't have to. You can copy your games from one sheet to the other. A few alterations have to be made tho.

Since I added a few more stats some columns have been changed. Before you copy your games and paste them into the new sheet, there is one thing you need to do. Go to your "General Stats" page and right click on column "X". Then, hit insert. This will add a new column between "KDA" and "CS/min". Then, you need to remove column V, right click it, hit delete. After that, select all the rows containing data from your games. Be sure to select the complete rows! (Click on the row number to select one cell, drag down to select multiple.) Then, copy these rows.  
After that you can go to your new spreadsheet and paste them in "General stats" in row **4**. **Leave row 3 empty.** Some formulas have been changed, new formulas have been added, and to not delete them, you have to leave this row empty.  
Then you're good to go! If you want to remove the empty row, you can do so after you've added at least one new game via the Update button. Then the formulas are stored in the cells of that new game, so you can remove the empty row. (Right click, remove.)  

**Changelog:**  
- v1.4:  
  - Added team side (SR)  
  - Added damage percentage dealt to champions from total team damage  
  - Added cs differences at 10, 20, 30 min  
  - Added cs/min for time intervals 0 - 10, 10 - 20, 20 - 30  
  - Added filters in the match history overview  
- v1.3:  
  - Added support for EUNE
  - Fixed bug with summoner names that contained spaces
  - Capital letters don't matter anymore in region (or summoner name)

****
### What's next
****
I'll be looking to improve the spreadsheet significantly. One of the most important areas in which it can be improved is speed. Currently it's pretty slow, I'll be looking to improve that in the next update. Also, I want to hear from you guys what stats or features you would like to see.  
Another feature I'd like to add in the future is the ability to add all your ranked games from this season. I haven't included it yet since it would take ages considering the speed that game insertion currently happens, but I hope to add it in the near future. 

****
### Oh no, I get an error!
****
That sucks :/ First, make sure you have entered your summoner name, region and API key correctly. If that is the case, contact me! When sending a PM to me ([on reddit](http://reddit.com/u/ikhebgeenaccount)) or opening an [issue on GitHub](https://github.com/ikhebgeenaccount/Excel-Match-History-Analyzer/issues), be sure to tell me what the error message is. Also, I'd really appreciate it if you could screenshot the page which is erroring and your front sheet.
