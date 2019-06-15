from datetime import date, timedelta, datetime
from geolocation.main import GoogleMaps
from dskyXLformat import dsky_format_xlbook
from console_tables import printTable
from darksky import forecast
import os, sys, subprocess
from numbar import *
import tkinter as tk
import argparse
import dictXL
import warnings


def open_file(filename):
    if sys.platform == "win32":
        os.startfile(filename)
    else:
        opener ="open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, filename])
        

##    home address lat and long
##    my_lat = 40.6359
##    my_lon = -77.5664

# darksky api key (free tier, 1000 requests/day)
key = '6a16aa916f1fe73085452ff8cfd5f8bb'
# Google Maps api key
google_maps = GoogleMaps(api_key='AIzaSyBe-Piv9VcT8gCXPp8bMEHdeUSgMxPV4Xw')

# setup command line argument parser and get args
parser = argparse.ArgumentParser()
parser.add_argument("-a", "--address", help="prompt for address", action='store_true')
args = parser.parse_args()

# look for -a or --address arguments from command line
# ask for a new address if found
if args.address:
    myaddress = input('Enter City, State: ')
    if not myaddress:
        sys.exit('No Address Entered')

    location = google_maps.search(location=myaddress) # sends search to Google Maps.
    my_location = location.first() # returns only first location.

    HOME=my_location.lat, my_location.lng
    newaddress = my_location.formatted_address
    
    mycounty = str(
        my_location.administrative_area[1].name
        ).lstrip("b'").split(' ')[0].lower()

    fstring = 'Forecast for {}'.format(newaddress)

else:
    address = 'Burnham N. Beech Street 208'
    try:
        location = google_maps.search(location=address) # sends search to Google Maps.
        my_location = location.first() # returns only first location.
        HOME = my_location.lat, my_location.lng

        myaddress = my_location.formatted_address#geo.get_address_fromll(my_lat,my_lon)

        mycounty = str(
            my_location.administrative_area[1].name
            ).lstrip("b'").split(' ')[0].lower()
    except:
        ##    home address lat and long
        my_lat = 40.6359
        my_lon = -77.5664
        HOME = my_lat, my_lon
        myaddress = 'Burnham, PA'
        mycounty = 'mifflin'

    fstring = 'Forecast for {}'.format(myaddress)
    
print(fstring)
# get the forecast and print some hourly and daily data
weekday = date.today()
with forecast(key, *HOME) as home:
    # weather summary for the week
    print(home.daily.summary, end='\n---\n')
    # iterate the days
    for day in home.daily:
        # load day dict
        day = dict(day = date.strftime(weekday, '%a'),
                   icon = day.icon,
                   daysum = day.summary,
                   tempMin = day.temperatureMin,
                   tempMax = day.temperatureMax,
                   ptype = day.precipType
                   if hasattr(day, 'precipType') else 'precip',
                   pcip = str(int(day.precipProbability * 100))
                   )
        # print day dict
        print('{day}: {daysum} Temp range: {tempMin} - {tempMax} | {ptype}: {pcip}%'.format(**day))
        # if we are on 'today' then print some hourly data
        if weekday==date.today():
            # iterate the hours
            for hour in home.hourly:
                # start with the current hour
                if datetime.fromtimestamp(hour.time).hour \
                   >= datetime.now().hour \
                   and datetime.fromtimestamp(hour.time).date() \
                   <= datetime.now().date():
                    # print some hourly data
                    # display the tempMax for the hour as a bar graph
                    # then print the icon name and precip probability
                    print(datetime.fromtimestamp(hour.time).strftime('%I %p'),
                          printNumBar(hour.temperature,
                                      day['tempMax'],
                                      length = 20),
                          '"' + hour.icon + '"',
                          '| ' + 'precip' + ': ' + str(int(hour.precipProbability*100))+'%')
        # increment date for display
        weekday += timedelta(days=1)

"""clear-day, clear-night, rain, snow, sleet, wind,
fog, cloudy, partly-cloudy-day, or partly-cloudy-night"""

# Insert home.daily data into daze list of dicts
daze = [
    {attr: var
     if 'time' not in attr.lower()
     else datetime.fromtimestamp(var)
     for attr, var in vars(day).items()
     if attr != '_data'
     } for day in home.daily
     ]

dazecols = [x for x in daze[0]]
a, b = dazecols.index('time'), 0
dazecols[b], dazecols[a] = dazecols[a], dazecols[b]

# Insert home.hourly data into hrly list of dicts
hrly=[
        {attr: var
         if 'time' not in attr.lower()
         else datetime.fromtimestamp(var)
         for attr, var in vars(m).items()
         if attr != '_data'
         } for m in home.hourly
         ]

hrlycols = [x for x in hrly[0]]
a, b = hrlycols.index('time'), 0
hrlycols[b], hrlycols[a] = hrlycols[a], hrlycols[b]

# If minutely data exists, insert home.minutely data into mints list of dicts
try:
    mints = [
        {attr: var
         if 'time' not in attr.lower()
         else datetime.fromtimestamp(var)
         for attr, var in vars(m).items()
         if attr != '_data'
         } for m in home.minutely
         ]

    mintcols = [x for x in mints[0]]
    a, b = mintcols.index('time'), 0
    mintcols[b], mintcols[a] = mintcols[a], mintcols[b]

except:
    mints = False
    
# Insert home.currently data into currently dict
currently = {l: v
             if 'time' not in str(l).lower()
             else datetime.fromtimestamp(v).strftime('%x %X')
             for l,v in vars(home.currently).items()
             if l != '_data'
             }

# prep for printing Current Conditions from currently dict
# Identify col widths for each "key + ': ' + value" combination, return the largest width
maxcol = max(
    [
        len(v + ': ' + str(currently[v]))
        for v in currently
        ]
        )

# Converts currently dict to 'list of lists'. Each individual list is
# a [key, value] from dict.
cy = [
    [l, str(v)
     if 'time' not in str(l).lower()
     else datetime.fromtimestamp(v).strftime('%x %X')
     ]
    for l,v
    in vars(home.currently).items()
    if l != '_data'
    ]

# if there are an odd number of lists ([key, value] pairs) add a blank one to the end
if len(cy) % 2 != 0:
    cy.append(['',''])

# creates a list of tuples (2), one containing the keys, the other containing the
# values then substitutes the length of each item for the item, then returns a list
# of the max value from each tuple [max, max]. These are col widths.
cz = [max(map(len,col)) for col in zip(*cy)]

# format string using the calculated col widths from above
# "label       : value         "
fs = ': '.join(["{{:<{}}}".format(i) for i in cz])

# apply format to the 'list of lists', effectivley flattening to a list of strings
# "label     : value     ", "label     : value     ",...
cy = [fs.format(*item) for item in cy]

# convert the flat list to 'list of lists' containing two items per list
# (col1 and col2) ["label     : value     ", "label     : value     "]
ml = [[x,y] for x,y in zip(cy[0::2],cy[1::2])]

# calculate the max col widths for the new 'list of lists'
cz = [max(map(len,col)) for col in zip(*ml)]

# format string using the calculated col widths from above
# "col 1 data    | col 2 data    "
fs = ' | '.join(["{{:<{}}}".format(i) for i in cz])

# print the list
# final result; "col 1 label  : col 1 value    | col 2 label   : col 2 value    "
print(); print('Current Conditions')
for item in ml: print(fs.format(*item))

xlfile = 'dsky.xlsx'
rxlfile = r'dsky.xlsx'

with warnings.catch_warnings():
    warnings.filterwarnings("ignore")

    # output mints dict (home.minutely data) to file
    if mints:
        dictXL.dict2xl(mints, xlfile, 'Minutely', mintcols)
        t=printTable(mints)
        f=open('mints.txt','w')
        f.write(t)
        f.close()

    # output hrly dict (home.hourly data) to file
    if hrly:
        dictXL.dict2xl(hrly, xlfile, 'Hourly', hrlycols)
        t=printTable(hrly)
        f=open('hours.txt','w')
        f.write(t)
        f.close()

    # output daze dict (home.daily data) to file
    dictXL.dict2xl(daze, xlfile, 'Daily', dazecols)
    t=printTable(daze)
    f=open('daze.txt','w')
    f.write(t)
    f.close()

# if there are any weather alerts for Mifflin County, display them in a pop-up window
try:
    alerts = [
        ('expires: ' + datetime.fromtimestamp(x.expires).strftime('%x %X') + '\n',
         x.description)
        for x in home.alerts
        if mycounty in str(x.regions).lower()
        ]
    
    if alerts:
        for a in alerts:
            root = tk.Tk()
            root.title(a[0])
            label = tk.Message(root, text=a[1], font=("Arial", 12))
            label.pack(side="top", fill="both", expand=True, padx=20, pady=5)
            f = tk.Frame(root, height=50, width=75)
            f.pack_propagate(0) # don't shrink
            f.pack()
            button = tk.Button(f, text="OK", command=lambda: root.destroy())
            button.pack(side="bottom", fill='both', expand=True, padx=5, pady=5)
            root.mainloop()
except AttributeError:
    print(); print('No Alerts')

dsky_format_xlbook(xlfile)
open_file(rxlfile)
