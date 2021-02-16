# nhl-fantasy-drafter
A VBA tool to assist with NHL fantasy draft choices. Jupyter notebook code pulls player stats from the NHL API for the last three seasons.

## Instructions:

### Extracting Data
* Run code in nhl_data_pull.ipynb to get NHL player stats for the last three seasons for current roster players. Currently the season years need to be manually updated at the start of the code.

### Using Tool
* In the 'Info' tab, use the "Get Data" button to get the stats from the last 3 seasons for current roster players. Note: the default spreadsheet uses 2019/2020, 2018/2019 & 2017/2018 data.
* Enter the point values for your league in the "Point Values" section of the 'Info' tab; if it's a categories league, use the proportion calculator to determine the value of each stat.
* Run the "Rank Skaters" macro in the 'Info' tab to determine the value of each player.
* During the draft, keep track of drafted players by highlighting their names in red (specifically vbRed) in the 'Rankings' tab. When it is your turn to pick, run the macro in the "RemainingPlayers" tab to filter for the top remaining skater selection.

