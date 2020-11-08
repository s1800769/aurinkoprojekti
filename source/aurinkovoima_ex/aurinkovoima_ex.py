import sys, os, json, datetime, requests, winreg
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from glob import glob
from datetime import timedelta
from svglib.svglib import svg2rlg
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.graphics import renderPDF
from reportlab.platypus import Table, TableStyle, Image
from reportlab.pdfbase.pdfmetrics import stringWidth

useTestWeatherData = False

#PDF-tiedoston nimi, PDF leveys, PDF pituus, API-avain, säädatan leveysaste, säädatan pituusaste, asetusten tallentaminen on/off, säädatan tunti
settings = ['aurinkovoima_ex_raportti', '595', '842', 'QeOdL1ZU2YQ2xX6nJbE9QFLHBW7wc5yt', '60.10', '24.56', 13]

def resourcePath(relativePath):
    try:
        basePath = sys._MEIPASS
    except:
        basePath = os.path.dirname(os.path.realpath(__file__)) + '/images'
    return os.path.join(basePath, relativePath)

with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as key:
    downloadsDirectory = winreg.QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]
try:
    latest = max(glob(downloadsDirectory + '/*.xlsx'), key = os.path.getctime)
except ValueError:
    sys.exit()

doc = canvas.Canvas(os.path.dirname(os.path.realpath(__file__)) + '/' + settings[0] + '.pdf')
docSize = {'width': int(settings[1]), 'height': int(settings[2])}

unit = (docSize['width'] + docSize['height']) / 2 * 0.015

def parseExcel(excel = latest):
    data = pd.read_excel(excel)
    columns = pd.DataFrame(data, columns = ['Päivämäärä ja aika', 'Energia invertteriä kohti|Symo 10.0-3-M (# 1)', 'Energia invertteriä kohti / kWp|Symo 10.0-3-M (# 1)', 'Järjestelmä yhteensä'])
    columns.rename(columns = {'Päivämäärä ja aika' : 0, 'Energia invertteriä kohti|Symo 10.0-3-M (# 1)' : 1, 'Energia invertteriä kohti / kWp|Symo 10.0-3-M (# 1)' : 2, 'Järjestelmä yhteensä' : 3}, inplace = True)
    columns.drop(columns.index[:1], inplace = True)
    return {'dates': list(columns[0]), 'kWhPerInverter': list(columns[1]), 'kWhBykWpPerInverter': list(columns[2]), 'kWhTotal': list(columns[3])}

def convertDates(dates, dateFormat):
    convertedDates = []
    for dateData in dates:
        dateValues = datetime.datetime.strptime(str(dateData), '%Y-%m-%d %H:%M:%S')
        convertedDates.append(dateFormat.format(second = dateValues.second, minute = dateValues.minute, hour = dateValues.hour, day = dateValues.day, month = dateValues.month, year = dateValues.year))
    return convertedDates

def getWeatherData(dates, weatherCoords = {'lat': settings[4], 'lon': settings[5]}, weatherHour = settings[6], key = settings[3]):
    dates.append(dates[-1] + timedelta(days = 1))
    cDates = convertDates(dates, '{year}-{month:02d}-{day:02d}T{hour:02d}:{minute:02d}:{second:02d}Z')
    timeframes = ((cDates[0], cDates[1]), (cDates[1], cDates[2]), (cDates[2], cDates[3]), (cDates[3], cDates[4]), (cDates[4], cDates[5]), (cDates[5], cDates[6]), (cDates[6], cDates[7]))
    datesData = []
    for index, val in enumerate(timeframes):
        date = dates[index]
        query = {'lat': weatherCoords['lat'], 'lon': weatherCoords['lon'], 'unit_system': 'si', 'start_time': val[0], 'end_time': val[1], 'fields': 'temp,cloud_cover,precipitation_type', 'apikey': key}
        response = requests.request('GET', 'https://api.climacell.co/v3/weather/historical/station', params = query)
        data = json.loads(response.text)
        temp = (str(data[weatherHour]['temp']['value']) + '°C')
        precipitation = data[weatherHour]['precipitation_type']['value']
        cloudCover = (data[weatherHour]['cloud_cover']['value'])
        if not cloudCover:
            cloudCover = 0
        datesData.append({'date': date, 'temp': temp, 'precipitation': precipitation, 'cloudCover': cloudCover})
    dates.pop(7)
    return datesData

def drawWeatherInfographic(data, topMargin, printCoords = False, doc = doc, docSize = docSize):
    datesDate = []
    datesDisplay = []
    datesTemp = []
    for item in data:
        datesDate.append(item['date'])
        datesTemp.append(item['temp'])
        if item['precipitation'] == 'rain':
            dateDisplay = resourcePath('rainy_cloud.png')
        elif item['precipitation'] == 'snow':
            dateDisplay = resourcePath('snowy_cloud.png')
        else:
            if item['cloudCover'] >= 50:
                dateDisplay = resourcePath('cloud.png')
            elif item['cloudCover'] >= 25:
                dateDisplay = resourcePath('cloudy.png')
            else:
                dateDisplay = resourcePath('sun.png')
        dateDisplay = Image(dateDisplay, (unit * 5), (unit * 5))
        datesDisplay.append(dateDisplay)
    datesDate = convertDates(datesDate, '{day:02d}/{month:02d}/{year}')
    weatherInfographic = Table([datesTemp, datesDisplay, datesDate])
    weatherInfographic.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('FONTSIZE', (0,0), (-1, 0), (unit * 1.25)), ('FONTSIZE', (0,2), (-1, 2), unit), ('TOPPADDING', (0,1), (-1, 1), unit), ('BOTTOMPADDING', (0,1), (-1, 1), unit)]))
    weatherInfographic.canv = doc
    infographicWidth, infographicHeight = weatherInfographic.wrap(0, 0)
    xAxis = (docSize['width'] - infographicWidth) / 2
    yAxis = docSize['height'] - infographicHeight - topMargin
    weatherInfographic.drawOn(doc, xAxis, yAxis)
    coords = [{'x': xAxis, 'y': docSize['height'] - yAxis}]
    if printCoords:
        print(coords)
    return coords

def drawText(text, topMargin = 0, leftMargin = 0, font = {'family': 'Helvetica', 'size': unit}, printCoords = False, doc = doc, docSize = docSize):
    xAxis = leftMargin
    if xAxis == 0:
        xAxis += (docSize['width'] - stringWidth(text, font['family'], font['size'])) / 2
    yAxis = docSize['height'] - font['size'] - topMargin
    textObject = doc.beginText(xAxis, yAxis)
    textObject.setFont(font['family'], font['size'])
    textObject.textOut(text)
    doc.drawText(textObject)
    coords = [{'x': xAxis, 'y': docSize['height'] - yAxis}]
    if printCoords:
        print(coords)
    return coords

def makeGraph(width, height, graphTitle, graphLabel, data, tickSpacing, tickRotation):
    graphFigure = plt.figure(figsize = (width, height))
    plt.bar(data[0], data[1], color = 'orange')
    plt.title(graphTitle, size = (unit * 0.75))
    plt.ylabel(graphLabel, size = (unit * 0.6))
    plt.yticks(size = (unit * 0.6))
    plt.xticks(size = (unit * 0.6), ticks = plt.xticks()[0][::tickSpacing], rotation = tickRotation)
    graphBuffer = BytesIO()
    graphFigure.savefig(graphBuffer, format = 'svg')
    graphBuffer.seek(0)
    image = svg2rlg(graphBuffer)
    return image

def drawGraphs(graphs, topMargin, spaceBetween, printCoords = False, doc = doc, docSize = docSize):
    imagesTotalWidth = 0
    for graph in graphs:
        imagesTotalWidth += graph.width
    if len(graphs) > 1:
        spaceAfterGraphs = spaceBetween * (len(graphs) - 1)
    else:
        spaceAfterGraphs = 0
    freeSpace = docSize['width'] - imagesTotalWidth - spaceAfterGraphs
    sidesMargin = freeSpace / 2
    if sidesMargin < 0:
        return False
    xAxis = sidesMargin
    coords = []
    for graph in graphs:
        yAxis = docSize['height'] - graph.height - topMargin
        renderPDF.draw(graph, doc, xAxis, yAxis)
        coords.append({'x': xAxis, 'y': docSize['height'] - yAxis})
        xAxis += graph.width + spaceBetween
    if printCoords:
        print(coords)
    return coords

try:
    doc.setPageSize((docSize['width'], docSize['height']))
    coords = drawText('Aurinkopaneelien energiatuotanto', (unit * 3), 0, {'family': 'Helvetica-Bold', 'size': (unit * 2)})
    excelData = parseExcel()
    dataRowLengthGreaterThanWeek = False
    for column in excelData:
        if len(excelData[column]) > 7:
            dataRowLengthGreaterThanWeek = True
            break
    convertedDates = convertDates(excelData['dates'], '{day:02d}/{month:02d}/{year}')
    doc.setTitle('Raportti ' + convertedDates[0] + '-' + convertedDates[-1])
    coords = drawText('Raportti ajalta: ' + convertedDates[0] + '-' + convertedDates[-1], coords[0]['y'] + (unit * 2), (unit * 4))
    if not dataRowLengthGreaterThanWeek:
        if useTestWeatherData:
            weatherData = [{'date': datetime.datetime(2020, 3, 30, 0, 0), 'temp': '-1°C', 'precipitation': 'none', 'cloudCover': 0}, {'date': datetime.datetime(2020, 3, 31, 0, 0), 'temp': '-1°C', 'precipitation': 'none', 'cloudCover': 0}, {'date': datetime.datetime(2020, 4, 1, 0, 0), 'temp': '4°C', 'precipitation': 'none', 'cloudCover': 0}, {'date': datetime.datetime(2020, 4, 2, 0, 0), 'temp': '5°C', 'precipitation': 'none', 'cloudCover': 0}, {'date': datetime.datetime(2020, 4, 3, 0, 0), 'temp': '2°C', 'precipitation': 'none', 'cloudCover': 0}, {'date': datetime.datetime(2020, 4, 4, 0, 0), 'temp': '6°C', 'precipitation': 'none', 'cloudCover': 0}, {'date': datetime.datetime(2020, 4, 5, 0, 0), 'temp': '6°C', 'precipitation': 'none', 'cloudCover': 0}]
        else:
            weatherData = getWeatherData(excelData['dates'])
        coords = drawWeatherInfographic(weatherData, coords[0]['y'] + (unit * 2))
    convertedDates = convertDates(excelData['dates'], '{day:02d}/{month:02d}')
    graphsData = ((convertedDates, excelData['kWhPerInverter']), (convertedDates, excelData['kWhBykWpPerInverter']), (convertedDates, excelData['kWhTotal']))
    tickSettings = [1, 0, 0]
    if dataRowLengthGreaterThanWeek:
        tickSettings = [4, 20, 0]
    coords = drawGraphs([makeGraph(((unit * 20) / 72), ((unit * 20) / 72), 'Energia invertteriä kohti kWh', 'kWh', graphsData[0], tickSettings[0], tickSettings[1]), makeGraph(((unit * 20) / 72), ((unit * 20) / 72), 'Energia invertteriä kohti kWh / kWp', 'kWh / kWp', graphsData[1], tickSettings[0], tickSettings[1])], coords[0]['y'] + (unit * 2), 0)
    drawGraphs([makeGraph(((unit * 40) / 72), ((unit * 20) / 72), 'Järjestelmä yhteensä', 'kWh', graphsData[2], tickSettings[0], tickSettings[2])], coords[0]['y'] + unit, 0)
    doc.showPage()
    doc.save()
    os.startfile(os.path.dirname(os.path.realpath(__file__)) + '/' + settings[0] + '.pdf')
except:
    sys.exit()