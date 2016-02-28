****
### What is it  
****
Match History Analyzer is an Excel Spreadsheet that takes data from your games and calculates stats such as your CS/min and Gold/min on a per game basis, but also over all games you have played. It shows you these stats for all your champions, but you can also filter for certain champions, even by role.  
It also calculates your winrate and playrate vs certain matchups, and allows you to add notes to these matchups. 

****
#### What do I need
****
Microsoft Office Excel installed, on Windows. Mac is not supported (yet), since it uses a different version of VBA than Windows. (sorry)

#### [Download link](https://github.com/ikhebgeenaccount/Excel-Match-History-Analyzer/releases/tag/1.3)

****
### Supporting and contributing
****
If you'd like to help with this project, contact me via reddit, [http://reddit.com/u/ikhebgeenaccount](/u/ikhebgeenaccount).

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
After you have updated the spreadsheet for the first time, you can update it after each game you play (or after multiple games) and it will insert those new games. Old games will not be deleted, so it takes some time to get useful data from the sheet.

****
### When you have downloaded the previous version
****
First of all, nice you downloaded it :) Now, you probably don't want to start over with collecting your games. Dont worry tho, you won't have to. You can select the rows of the games you have in your previous sheet and copy them to the new sheet.

**Changelog:**  
- v1.4:  
  - Added team side (SR)  
  - Added damage percentage dealt to champions from total team damage  
  - Added cs differences at 10, 20, 30 min  
  - Added cs/min for time intervals 0 - 10, 10 - 20, 20 - 30  
- v1.3:  
  - Added support for EUNE
  - Fixed bug with summoner names that contained spaces
  - Capital letters don't matter anymore in region (or summoner name)

****
### What's next
****
I'll be looking to improve the spreadsheet significantly. One of the most important areas in which it can be improved is speed. Currently it's pretty slow, I'll be looking to improve that in the next update. Also, I want to hear from you guys what stats or features you would like to see.  
Another feature I'd like to add in the future is the ability to add all your ranked games from this season. I haven't included it in this update since it would take ages considering the speed that game insertion currently happens, but I hope to add it in the near future. 

****
### Oh no, I get an error!
****
That sucks :/ First, make sure you have entered your summoner name, region and API key correctly. If that is the case, contact me! When sending a PM to me, commenting here or opening an [issue on GitHub](https://github.com/ikhebgeenaccount/Excel-Match-History-Analyzer/issues), be sure to tell me what the error message is. Also, I'd really appreciate it if you could screenshot the page which is erroring.
