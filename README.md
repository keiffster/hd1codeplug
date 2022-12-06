# hd1codeplug
This is an excel worksheet and associated Python script that is used to create the neccassary files for an HD1 Code Pug

## Overview
I have just recently entered the world of Amatuer Radio and one of the first radios I puchased was the Ailunce HD1 DMR Radio. Anyone who is involved in DMR knows there is a trade off between price and ability to programme. Also programme requires you to get your head round Talkgroups and Channels and Contacts. I started with a simple Excel spreadsheet to track different talkgroups and channels I wanted to use. I also started playing around with Hotspots given I live in a pretty littel village with limited access to repeaters. I ended up with 3 hotposts, one for DMR+, one fo Brandmeister and one for TGIF having repurposed a bunch of Raspberry PI's I had lying round.

As I got more and more into DMR and using the Talkgroups I started playing around with automating the creating of the neccassayr CSV files that the HD1 CPS uses. Python is my goto scripting langauge and this code is the result of me turning my ideas into a usable script. 

## Installation 
The code itself relies on the open 'openpyxl' Python library. To install use pip

pip install openpyxl

The copy the scipt into the folder that contains you excel spreasheet

## Create an Excel spreadsheet
python3 codeplug.py input.xlsx xlsx output.xlsx

## Create associated CSV Files
python3 codeplug.py input.xlsx csv

This will create the following files that can be loaded into your HD1 software

- HD1 Channel Information.csv
- HD1 Priority Contacts.csv
- HD1 Address Book Contacts.csv

## Template Excel Sheet

- HD! Base Info
- HD1 Priority Contacts
- HD1 Channel Information
- HD1 Zone Information
- HD1 Address Book Contacts
- DVS Talkgroups
- BM Talkgroups
- TGIF Talkgroups
- VFO Channel Info
- PMR Channel Info
- Marine Channel Info
- UK Analog Repeaters
- UK Digital DMR Repeaters



