import sys, os, json, datetime, requests
import pandas as pd
import tkinter as tk
import matplotlib.pyplot as plt
import tkinter.filedialog
from io import BytesIO
from datetime import timedelta
from svglib.svglib import svg2rlg
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.graphics import renderPDF
from reportlab.platypus import Table, TableStyle, Image
from reportlab.pdfbase.pdfmetrics import stringWidth

useTestWeatherData = False

def resourcePath(relativePath):
    try:
        basePath = sys._MEIPASS
    except:
        basePath = os.path.dirname(os.path.realpath(__file__)) + '/images'
    return os.path.join(basePath, relativePath)

window = tk.Tk()
window.title('Aurinkovoima')

#Func quitProgram
def quitProgram():
    window.quit()
    window.destroy()

window.protocol('WM_DELETE_WINDOW', quitProgram)

#mainPanel, settingsPanel
panelSizes = [{'width': 300, 'height': 300}, {'width': 150, 'height': 300}]
panelPaddings = [{'x': 20, 'y': 10}, {'x': 20, 'y': 10}]
panelBackgrounds = ['#BFBFBF', '#DBDBDB']

#mainTitle, mainSub, mainResponse
mainTexts = ['Aurinkovoima', 'Muuttaa Excel-tiedoston aurinkopaneelidatan helposti luettavaksi PDF-tiedostoksi.', '']
mainTextFonts = [('Helvetica', 16, 'bold'), ('Helvetica', 10), ('Helvetica', 10)]
mainTextPaddings = [{'x': 0, 'y': 0}, {'x': 0, 'y': 0},  {'x': 0, 'y': 0}]
mainTextMargins = [{'x': 0, 'y': 0}, {'x': 0, 'y': 3}, {'x': 0, 'y': 3}]

#mainFileSelect
mainFileSelectText = 'Valitse tiedosto'
mainFileSelectPadding = {'x': 5, 'y': 5}
mainFileSelectMargin = {'x': 0, 'y': 5}

#settingsTitle, settingsDocNameSub, settingsPageSizeSub, settingsApiKeySub, settingsWeatherCoordsSub
settingsTexts = ['Asetukset', 'PDF-tuotoksen nimi:', 'Sivun koko:', 'ClimaCell API-avain:', 'Säädatan koordinaatit:']
settingsTextFonts = [('Helvetica', 12, 'bold'), ('Helvetica', 10), ('Helvetica', 10), ('Helvetica', 10), ('Helvetica', 10)]
settingsTextPaddings = [{'x': 0, 'y': 0}, {'x': 0, 'y': 0}, {'x': 0, 'y': 0}, {'x': 0, 'y': 0}, {'x': 0, 'y': 0}]
settingsTextMargins = [{'x': 0, 'y': 4}, {'x': 0, 'y': 0}, {'x': 0, 'y': 0}, {'x': 0, 'y': 0}, {'x': 0, 'y': 0}]

#settingsPageSizeButtonA, settingsPageSizeButtonB, settingsPageSizeButtonC
settingsPageSizeButtonTexts = ['A4 (595pt x 842pt)', 'A5 (420pt x 595pt)', 'Muu']
settingsPageSizeButtonPaddings = [{'x': 0, 'y': 0}, {'x': 0, 'y': 0}, {'x': 0, 'y': 0}]
settingsPageSizeButtonMargins = [{'x': 0, 'y': 0}, {'x': 0, 'y': 0}, {'x': 0, 'y': 0}]
settingsPageSizeVar = tk.IntVar()
settingsPageSizeButtonValues = [1, 2, 3]

#pageSizeCustomWidthEntry, pageSizeCustomHeightEntry
pageSizeCustomEntryFonts = [('Helvetica', 10), ('Helvetica', 10)]
pageSizeCustomEntryWidths = [6, 6]
pageSizeCustomEntryMargins = [{'x': 0, 'y': 0}, {'x': 6, 'y': 0}]

#settingsPrintWeatherData, settingsSaveConfig
settingsCheckboxTexts = ['Näytä säätaulukko', 'Tallenna asetukset']
settingsCheckboxFonts = [('Helvetica', 10), ('Helvetica', 10)]
settingsCheckboxPaddings = [{'x': 0, 'y': 0}, {'x': 0, 'y': 0}]
settingsCheckboxMargins = [{'x': 0, 'y': 4}, {'x': 0, 'y': 4}]
settingsCheckboxVars = [tk.IntVar(), tk.IntVar()]

#settingsDocNameEntry, settingsApiKeyEntry
settingsEntryFonts = [('Helvetica', 10), ('Helvetica', 10)]
settingsEntryWidths = [18, 18]
settingsEntryMargins = [{'x': 0, 'y': 0}, {'x': 0, 'y': 0}]

#weatherCoordsLatEntry, weatherCoordsLonEntry
weatherEntryFonts = [('Helvetica', 10), ('Helvetica', 10)]
weatherEntryWidths = [6, 6]
weatherEntryMargins = [{'x': 0, 'y': 0}, {'x': 6, 'y': 0}]

#master = window
mainPanel = tk.Frame(master = window, width = panelSizes[0]['width'], height = panelSizes[0]['height'], padx = panelPaddings[0]['x'], pady = panelPaddings[0]['y'], bg = panelBackgrounds[0])
settingsPanel = tk.Frame(master = window, width = panelSizes[1]['width'], height = panelSizes[1]['height'], padx = panelPaddings[1]['x'], pady = panelPaddings[1]['y'], bg = panelBackgrounds[1])

#master = mainPanel
mainTitle = tk.Label(master = mainPanel, text = mainTexts[0], font = mainTextFonts[0], padx = mainTextPaddings[0]['x'], pady = mainTextPaddings[0]['y'], bg = panelBackgrounds[0], wraplength = panelSizes[0]['width'])
mainSub = tk.Label(master = mainPanel, text = mainTexts[1], font = mainTextFonts[1], padx = mainTextPaddings[1]['x'], pady = mainTextPaddings[1]['y'], bg = panelBackgrounds[0], wraplength = panelSizes[0]['width'])
mainResponse = tk.Label(master = mainPanel, text = mainTexts[2], font = mainTextFonts[2], padx = mainTextPaddings[2]['x'], pady = mainTextPaddings[2]['y'], bg = panelBackgrounds[0], wraplength = panelSizes[0]['width'])

#master = settingsPanel
settingsTitle = tk.Label(master = settingsPanel, text = settingsTexts[0], font = settingsTextFonts[0], padx = settingsTextPaddings[0]['x'], pady = settingsTextPaddings[0]['y'], bg = panelBackgrounds[1], wraplength = panelSizes[1]['width'])
settingsDocNameSub = tk.Label(master = settingsPanel, text = settingsTexts[1], font = settingsTextFonts[1], padx = settingsTextPaddings[1]['x'], pady = settingsTextPaddings[1]['y'], bg = panelBackgrounds[1], wraplength = panelSizes[1]['width'], anchor = 'w')
settingsDocNameEntry = tk.Entry(master = settingsPanel, font = settingsEntryFonts[0], width = settingsEntryWidths[0])
settingsPageSizeSub = tk.Label(master = settingsPanel, text = settingsTexts[2], font = settingsTextFonts[2], padx = settingsTextPaddings[1]['x'], pady = settingsTextPaddings[2]['y'], bg = panelBackgrounds[1], wraplength = panelSizes[1]['width'], anchor = 'w')
settingsPageSizeButtonA = tk.Radiobutton(master = settingsPanel, text = settingsPageSizeButtonTexts[0], padx = settingsPageSizeButtonPaddings[0]['x'], pady = settingsPageSizeButtonPaddings[0]['y'], bg = panelBackgrounds[1], variable = settingsPageSizeVar, value = settingsPageSizeButtonValues[0], anchor = 'w')
settingsPageSizeButtonB = tk.Radiobutton(master = settingsPanel, text = settingsPageSizeButtonTexts[1], padx = settingsPageSizeButtonPaddings[1]['x'], pady = settingsPageSizeButtonPaddings[1]['y'], bg = panelBackgrounds[1], variable = settingsPageSizeVar, value = settingsPageSizeButtonValues[1], anchor = 'w')
settingsPageSizeButtonC = tk.Radiobutton(master = settingsPanel, text = settingsPageSizeButtonTexts[2], padx = settingsPageSizeButtonPaddings[2]['x'], pady = settingsPageSizeButtonPaddings[2]['y'], bg = panelBackgrounds[1], variable = settingsPageSizeVar, value = settingsPageSizeButtonValues[2], anchor = 'w')
settingsPageSizeCustomFrame = tk.Frame(master = settingsPanel, bg = panelBackgrounds[1])
settingsPrintWeatherData = tk.Checkbutton(master = settingsPanel, text = settingsCheckboxTexts[0], font = settingsCheckboxFonts[0], padx = settingsCheckboxPaddings[0]['x'], pady = settingsCheckboxPaddings[0]['y'], bg = panelBackgrounds[1], variable = settingsCheckboxVars[0], anchor = 'w')
settingsApiKeySub = tk.Label(master = settingsPanel, text = settingsTexts[3], font = settingsTextFonts[3], padx = settingsTextPaddings[3]['x'], pady = settingsTextPaddings[3]['y'], bg = panelBackgrounds[1], wraplength = panelSizes[1]['width'], anchor = 'w')
settingsApiKeyEntry = tk.Entry(master = settingsPanel, font = settingsEntryFonts[1], width = settingsEntryWidths[1])
settingsWeatherCoordsSub = tk.Label(master = settingsPanel, text = settingsTexts[4], font = settingsTextFonts[4], padx = settingsTextPaddings[4]['x'], pady = settingsTextPaddings[4]['y'], bg = panelBackgrounds[1], wraplength = panelSizes[1]['width'], anchor = 'w')
settingsWeatherCoordsFrame = tk.Frame(master = settingsPanel, bg = panelBackgrounds[1])
settingsSaveConfig = tk.Checkbutton(master = settingsPanel, text = settingsCheckboxTexts[1], font = settingsCheckboxFonts[1], padx = settingsCheckboxPaddings[1]['x'], pady = settingsCheckboxPaddings[1]['y'], bg = panelBackgrounds[1], variable = settingsCheckboxVars[1], anchor = 'w')

#master = settingsPageSizeCustomFrame
pageSizeCustomWidthEntry = tk.Entry(master = settingsPageSizeCustomFrame, font = pageSizeCustomEntryFonts[0], width = pageSizeCustomEntryWidths[0])
pageSizeCustomHeightEntry = tk.Entry(master = settingsPageSizeCustomFrame, font = pageSizeCustomEntryFonts[1], width = pageSizeCustomEntryWidths[1])

#master = settingsWeatherCoordsFrame
weatherCoordsLatEntry = tk.Entry(master = settingsWeatherCoordsFrame, font = weatherEntryFonts[0], width = weatherEntryWidths[0])
weatherCoordsLonEntry = tk.Entry(master = settingsWeatherCoordsFrame, font = weatherEntryFonts[1], width = weatherEntryWidths[1])

#PDF-tiedoston nimi, sivun-koko valinta, custom leveys, custom pituus, säädatan tulostus on/off, API-avain, säädatan leveysaste, säädatan pituusaste, asetusten tallentaminen on/off, säädatan tunti
defaultSettings = ['aurinkovoima_raportti', 1, '595', '842', 1, 'QeOdL1ZU2YQ2xX6nJbE9QFLHBW7wc5yt', '60.10', '24.56', 0, 13]

docName = defaultSettings[0]
pageSizeSelection = defaultSettings[1]
pageSizeCustomWidth = defaultSettings[2]
pageSizeCustomHeight = defaultSettings[3]
printWeather = defaultSettings[4]
key = defaultSettings[5]
weatherLat = defaultSettings[6]
weatherLon = defaultSettings[7]
saveConfig = defaultSettings[8]

if os.path.isfile(os.path.dirname(os.path.realpath(__file__)) + '/aurinkovoima_config.json'):
    with open(os.path.dirname(os.path.realpath(__file__)) + '/aurinkovoima_config.json') as settingsFile:
        settingsFileData = json.load(settingsFile)
        if settingsFileData['config']['docName']:
            docName = settingsFileData['config']['docName']
        if settingsFileData['config']['pageSizeSelection']:
            pageSizeSelection = settingsFileData['config']['pageSizeSelection']
        if settingsFileData['config']['pageSizeCustomWidth']:
            pageSizeCustomWidth = settingsFileData['config']['pageSizeCustomWidth']
        if settingsFileData['config']['pageSizeCustomHeight']:
            pageSizeCustomHeight = settingsFileData['config']['pageSizeCustomHeight']
        if settingsFileData['config']['printWeather']:
            printWeather = settingsFileData['config']['printWeather']
        elif settingsFileData['config']['printWeather'] == 0:
            printWeather = settingsFileData['config']['printWeather']
        if settingsFileData['config']['key']:
            key = settingsFileData['config']['key']
        if settingsFileData['config']['weatherLat']:
            weatherLat = settingsFileData['config']['weatherLat']
        if settingsFileData['config']['weatherLon']:
            weatherLon = settingsFileData['config']['weatherLon']
        if settingsFileData['config']['saveConfig']:
            saveConfig = settingsFileData['config']['saveConfig']

settingsDocNameEntry.insert(0, docName)
settingsPageSizeVar.set(pageSizeSelection)
if pageSizeSelection == 3:
    pageSizeCustomWidthEntry.insert(0, pageSizeCustomWidth)
    pageSizeCustomHeightEntry.insert(0, pageSizeCustomHeight)
settingsCheckboxVars[0].set(printWeather)
settingsApiKeyEntry.insert(0, key)
weatherCoordsLatEntry.insert(0, weatherLat)
weatherCoordsLonEntry.insert(0, weatherLon)
settingsCheckboxVars[1].set(saveConfig)

#Func mainFileSelect
def selectFile():
    excel = tkinter.filedialog.askopenfile(mode = 'r', initialdir = os.path.dirname(os.path.realpath(__file__)) + '/', title = 'Select File', filetypes = [('Excel', '*.xlsx'), ('All Files', '*')])

    if excel:
        #settingsDocNameEntry, settingsPageSizeButtonA/settingsPageSizeButtonB/settingsPageSizeButtonC, pageSizeCustomWidthEntry, pageSizeCustomHeightEntry, settingsPrintWeatherData, settingsApiKeyEntry, weatherCoordsLatEntry, weatherCoordsLonEntry, settingsSaveConfig
        settings = [settingsDocNameEntry.get(), settingsPageSizeVar.get(), pageSizeCustomWidthEntry.get(), pageSizeCustomHeightEntry.get(), settingsCheckboxVars[0].get(), settingsApiKeyEntry.get(), weatherCoordsLatEntry.get(), weatherCoordsLonEntry.get(), settingsCheckboxVars[1].get()]

        for index, value in enumerate(settings):
            if index == 0 or index == 5 or index == 6 or index == 7:
                if not value:
                    settings[index] = defaultSettings[index]
            if index == 1 and value == 3:
                for x in range(2):
                    try:
                        int(settings[index + (x + 1)])
                    except ValueError:
                        settings[index + x] = defaultSettings[index + (x + 1)]
            if index == 6 or index == 7:
                try:
                    float(value)
                except ValueError:
                    settings[index] = defaultSettings[index]
                if len(value.rsplit('.')[-1]) != 2:
                    settings[index] = defaultSettings[index]

        doc = canvas.Canvas(os.path.dirname(os.path.realpath(__file__)) + '/' + settings[0] + '.pdf')

        if settings[1] == 1:
            docSize = {'width': 595, 'height': 842}
        elif settings[1] == 2:
            docSize = {'width': 420, 'height': 595}
        elif settings[1] == 3:
            docSize = {'width': int(settings[2]), 'height': int(settings[3])}

        #Yksikkö, jota käytetään määrittelemään PDF-tiedoston eri elementtien koot responsiivisesti
        unit = (docSize['width'] + docSize['height']) / 2 * 0.015

        def parseExcel(excel = excel):
            data = pd.read_excel(excel.name)
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

        def getWeatherData(dates, weatherCoords = {'lat': settings[6], 'lon': settings[7]}, weatherHour = defaultSettings[9], key = settings[5]):
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
            if settings[4] and not dataRowLengthGreaterThanWeek:
                if useTestWeatherData:
                    weatherData = [{'date': datetime.datetime(2020, 3, 30, 0, 0), 'temp': '-1°C', 'precipitation': 'none', 'cloudCover': 0}, {'date': datetime.datetime(2020, 3, 31, 0, 0), 'temp': '-1°C', 'precipitation': 'none', 'cloudCover': 0}, {'date': datetime.datetime(2020, 4, 1, 0, 0), 'temp': '4°C', 'precipitation': 'none', 'cloudCover': 0}, {'date': datetime.datetime(2020, 4, 2, 0, 0), 'temp': '5°C', 'precipitation': 'none', 'cloudCover': 0}, {'date': datetime.datetime(2020, 4, 3, 0, 0), 'temp': '2°C', 'precipitation': 'none', 'cloudCover': 0}, {'date': datetime.datetime(2020, 4, 4, 0, 0), 'temp': '6°C', 'precipitation': 'none', 'cloudCover': 0}, {'date': datetime.datetime(2020, 4, 5, 0, 0), 'temp': '6°C', 'precipitation': 'none', 'cloudCover': 0}]
                else:
                    weatherData = getWeatherData(excelData['dates'])
                coords = drawWeatherInfographic(weatherData, coords[0]['y'] + (unit * 2))
            convertedDates = convertDates(excelData['dates'], '{day:02d}/{month:02d}')
            graphsData = ((convertedDates, excelData['kWhPerInverter']), (convertedDates, excelData['kWhBykWpPerInverter']), (convertedDates, excelData['kWhTotal']))
            tickSettings = [1, 0, 0]
            responseText = 'Tiedosto ' + settings[0] + '.pdf luotu onnistuneesti.'
            if dataRowLengthGreaterThanWeek:
                tickSettings = [4, 20, 0]
                if settings[4]:
                    responseText = 'Säätaulukon tulostus poistettu käytöstä viikkoraporttia suurempaa raporttia tehdessä.\n\nTiedosto ' + settings[0] + '.pdf luotu onnistuneesti.'
            coords = drawGraphs([makeGraph(((unit * 20) / 72), ((unit * 20) / 72), 'Energia invertteriä kohti kWh', 'kWh', graphsData[0], tickSettings[0], tickSettings[1]), makeGraph(((unit * 20) / 72), ((unit * 20) / 72), 'Energia invertteriä kohti kWh / kWp', 'kWh / kWp', graphsData[1], tickSettings[0], tickSettings[1])], coords[0]['y'] + (unit * 2), 0)
            drawGraphs([makeGraph(((unit * 40) / 72), ((unit * 20) / 72), 'Järjestelmä yhteensä', 'kWh', graphsData[2], tickSettings[0], tickSettings[2])], coords[0]['y'] + unit, 0)
            doc.showPage()
            doc.save()
            if settings[8]:
                configContents = {
                    'config': {
                        'docName': settings[0],
                        'pageSizeSelection': settings[1],
                        'pageSizeCustomWidth': str(docSize['width']),
                        'pageSizeCustomHeight': str(docSize['height']),
                        'printWeather': settings[4],
                        'key': settings[5],
                        'weatherLat': settings[6],
                        'weatherLon': settings[7],
                        'saveConfig': settings[8]
                    }
                }
                with open(os.path.dirname(os.path.realpath(__file__)) + '/aurinkovoima_config.json', 'w') as configOutput:
                    json.dump(configContents, configOutput)
            mainResponse['text'] = responseText
            mainResponse.configure(fg = '#1E6600')
            os.startfile(os.path.dirname(os.path.realpath(__file__)) + '/' + settings[0] + '.pdf')
        except:
            mainResponse['text'] = 'PDF-tiedostoa luotaessa tuli virhe.'
            mainResponse.configure(fg = '#8C0000')
    else:
        mainResponse['text'] = ''

mainFileSelect = tk.Button(master = mainPanel, text = mainFileSelectText, padx = mainFileSelectPadding['x'], pady = mainFileSelectPadding['y'], command = selectFile)

#Pack master = settingsPageSizeCustomFrame
pageSizeCustomWidthEntry.pack(side = tk.LEFT, padx = pageSizeCustomEntryMargins[0]['x'], pady = pageSizeCustomEntryMargins[0]['y'])
pageSizeCustomHeightEntry.pack(side = tk.LEFT, padx = pageSizeCustomEntryMargins[1]['x'], pady = pageSizeCustomEntryMargins[1]['y'])

#Pack master = settingsWeatherCoordsFrame
weatherCoordsLatEntry.pack(side = tk.LEFT, padx = weatherEntryMargins[0]['x'], pady = weatherEntryMargins[0]['y'])
weatherCoordsLonEntry.pack(side = tk.LEFT, padx = weatherEntryMargins[1]['x'], pady = weatherEntryMargins[1]['y'])

#Pack master = mainPanel
mainTitle.pack(padx = mainTextMargins[0]['x'], pady = mainTextMargins[0]['y'])
mainSub.pack(padx = mainTextMargins[1]['x'], pady = mainTextMargins[1]['y'])
mainFileSelect.pack(padx = mainFileSelectMargin['x'], pady = mainFileSelectMargin['y'])
mainResponse.pack(padx = mainTextMargins[2]['x'], pady = mainTextMargins[2]['y'])

#Pack master = settingsPanel
settingsTitle.pack(padx = settingsTextMargins[0]['x'], pady = settingsTextMargins[0]['y'])
settingsDocNameSub.pack(fill = tk.BOTH, padx = settingsTextMargins[1]['x'], pady = settingsTextMargins[1]['y'])
settingsDocNameEntry.pack(padx = settingsEntryMargins[0]['x'], pady = settingsEntryMargins[0]['y'], anchor = 'w')
settingsPageSizeSub.pack(fill = tk.BOTH, padx = settingsTextMargins[2]['x'], pady = settingsTextMargins[2]['y'])
settingsPageSizeButtonA.pack(fill = tk.BOTH, padx = settingsPageSizeButtonMargins[0]['x'], pady = settingsPageSizeButtonMargins[0]['y'])
settingsPageSizeButtonB.pack(fill = tk.BOTH, padx = settingsPageSizeButtonMargins[1]['x'], pady = settingsPageSizeButtonMargins[1]['y'])
settingsPageSizeButtonC.pack(fill = tk.BOTH, padx = settingsPageSizeButtonMargins[2]['x'], pady = settingsPageSizeButtonMargins[2]['y'])
settingsPageSizeCustomFrame.pack(fill = tk.BOTH)
settingsPrintWeatherData.pack(fill = tk.BOTH, padx = settingsCheckboxMargins[0]['x'], pady = settingsCheckboxMargins[0]['y'])
settingsApiKeySub.pack(fill = tk.BOTH, padx = settingsTextMargins[3]['x'], pady = settingsTextMargins[3]['y'])
settingsApiKeyEntry.pack(padx = settingsEntryMargins[1]['x'], pady = settingsEntryMargins[1]['y'], anchor = 'w')
settingsWeatherCoordsSub.pack(fill = tk.BOTH, padx = settingsTextMargins[4]['x'], pady = settingsTextMargins[4]['y'])
settingsWeatherCoordsFrame.pack(fill = tk.BOTH)
settingsSaveConfig.pack(fill = tk.BOTH, padx = settingsCheckboxMargins[1]['x'], pady = settingsCheckboxMargins[1]['y'])

#Pack master = window
mainPanel.pack(fill = tk.BOTH, side = tk.LEFT, expand = True)
settingsPanel.pack(fill = tk.BOTH, side = tk.LEFT, expand = True)

window.mainloop()