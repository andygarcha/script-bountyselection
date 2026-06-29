# script-bountyselection
Selects bounties from google sheets for the CE community.

# Hi Brooks.
I made this script. You are super duper welcome. 

## Too long and I hate you
Do the following:
1. Download/clone this repository
2. Open a terminal
3. `cd` to this directory. For example, if this is in your `Downloads` folder, run `cd Downloads` and then `cd script-bountyselection`. You can hit tab along the way to give you hints.
4. Run the following:
```bash
# Windows
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python main.py all
```
```bash
# Mac
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python main.py all
```

## I care about our friendship Andy
Well thank you brooks! To reduce redundant issues, I've broken this down into multiple parts. Here's how this all works:
1. `pull`. This will pull the `.xlsx` files from the 'Potentials' and 'Retro Potentials' sheets, and saves them to `potentials.xlsx` and `retro.xlsx`. This way:
    - You aren't waiting on a multi-MB-sized request *every time*
    - After the first pull, you don't have to be internet connected to do this.
2. `pullce`. This will pull all of the uncleared objectives from the site and storing their information in `uncleareds.json`.
3. `extract`. This will pull all of the information out of `potentials.xlsx` and `retro.xlsx`, combine it with the information from the uncleareds you just pulled down, and store it all in `totalinfo.json`.
4. `select`. This is the randomness part. It pulls from the newly created `totalinfo.json` file, and writes to `selection.json`. It also 
> Take a look at the constants near the top of `main.py` if you would like to adjust how many games the script pulls per category - like if you wanted to do 5 steam potentials instead of 7.

> Also, if you'd like to change *how much each bounty is worth*, you can similarly adjust the constants near the top of `main.py`.
5. `selectdata`. This pulls down price information from the steam games. I probably could have left this in `select`, but since it was internet related, I thought it best to separate this out.
6. `output`. This writes the information about the selected bounties to a bunch of `.txt` files. This *was* all in a nice `.json` file, but this is just easier for you, and the `.json` file wasn't being nice with line breaks.