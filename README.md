# hd1codeplug
This is an excel worksheet and associated Python script that is used to create the neccassary files for an HD1 Code Pug

## Overview
I have just recently entered the world of Amatuer Radio and one of the first radios I puchased was the Ailunce HD1 DMR Radio. Anyone who is involved in DMR knows there is a trade off between price and ability to programme. Also programme requires you to get your head round Talkgroups and Channels and Contacts. I started with a simple Excel spreadsheet to track different talkgroups and channels I wanted to use. I also started playing around with Hotspots given I live in a pretty littel village with limited access to repeaters. I ended up with 3 hotposts, one for DMR+, one fo Brandmeister and one for TGIF having repurposed a bunch of Raspberry PI's I had lying round.

As I got more and more into DMR and using the Talkgroups I started playing around with automating the creating of the neccassayr CSV files that the HD1 CPS uses. Python is my goto scripting langauge and this code is the result of me turning my ideas into a usable script. 

## How It Works
Each DMR network uses talkgroups, they are all number 1 to NNNN, some are shared, others are not. Within the HD1 you cerate priority contacts for each talkgroup. You then create contacts for each talkgroup and for each hotspot you are using. This scripts saves a lot of the grunt work by
- Combining all the talkgroups into none repeating list of contacts with generic names sichas TG1, TG2, TG 3250 etc
- Creates a set of Channels for each network with the Alias you provide and assigns the appropriate priority contact
- Creates corresponding CSV files you can load into the HD1
- Creates an additional zone file you can use to create zones with the approriate contacts in your HD1

## Installation 
The code itself relies on the open 'openpyxl' Python library. To install use pip

pip install openpyxl

The copy the scipt into the folder that contains you excel spreasheet

## Create an Excel spreadsheet
python3 codeplug.py <input xlsx filename> xlsx <output xlsx filename>

e.g. python3 codeplug.py input.xlsx xlsx output.xlsx

## Create associated CSV Files
python3 codeplug.py <input xlsx filename> csv

e.g. python3 codeplug.py input.xlsx csv

This will create the following files that can be loaded into your HD1 software

- HD1 Channel Information.csv
- HD1 Priority Contacts.csv
- HD1 Address Book Contacts.csv

## Template Excel Sheet
The script current ships with a template xlsx file that should get you started. Its is Scotland focused and is set up to use the 3 hotspots and DMR networks that I play with

- HD1 Base Info. 
  - This contains the se4t of sheets that control the script
  - Before using it you need to set your Radio ID in Cell A2
  -  You then decide which set of talk groups and which networks you want to combine

- HD1 Priority Contacts
  - This is auto generated
  - This contains the list of all the priority contacts to load into your HD1

- HD1 Channel Information
  - This is auto generated
  - This contains the list of all the channels to load into your HD1

- HD1 Zone Information
  - This is auto generated
  - To help you set up your zones because the HD1 does not allow you to inport
  - It lists the channels it auto created and their number. You can then create which ever zones you want. I typically create one zone for each hotspot or channel grouping

- HD1 Address Book Contacts
  - This is optional, but a good place to keep all your contacts
  - If it exists an associated CSV file for importing to the HD1 will be created

- DVS Talkgroups
 - A list of DV Scotland talkgroups I typically use

- BM Talkgroups
 - A list of Brandmeister talkgroups I typically use

- TGIF Talkgroups
 - A list of TGIF talkgroups I typically use

- VFO Channel Info
  - The 2 VFO channels that are the first 2 channles in the channels sheet

- PMR Channel Info
  - List of UK PMR Channels

- Marine Channel Info
  - List of UK Marine channels ( Still work in progress )

- UK Analog Repeaters
  - List of UK Analog repeaters

- UK Digital DMR Repeaters
  - List of UK Digital repeaters


## TODO
The following is just a list of things I want to add to the script

- Add support for Marine Frequencies
- Make frequency validating config driven
- More documentation
- More error handling
- Installtion package

