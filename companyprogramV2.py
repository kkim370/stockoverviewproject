from datetime import date
from openpyxl import Workbook
from openpyxl.styles.numbers import FORMAT_PERCENTAGE_00
import re
import sys
import aiohttp
import asyncio
from bs4 import BeautifulSoup

async def titleStats(ticker, session) -> tuple[str]:
    url: str = f'https://stockanalysis.com/stocks/{ticker}/'.format(ticker)
    async with session.get(url) as resp:
        htmlRaw: str = str(await resp.read())
        parse: BeautifulSoup = BeautifulSoup(htmlRaw, 'html.parser')

        parsePrice: str = str(parse.find(name='div', attrs={'class':'text-4xl'}).getText())
        parsePE: str = str(parse.find_all(name='td', attrs={'class': 'whitespace-nowrap px-0.5 py-[1px] text-left text-smaller font-semibold tiny:text-base xs:px-1 sm:py-2 sm:text-right sm:text-small'})[5].getText())
        yDivYield: str = str(parse.find_all(name='td', attrs={'class': 'whitespace-nowrap px-0.5 py-[1px] text-left text-smaller font-semibold tiny:text-base xs:px-1 sm:py-2 sm:text-right sm:text-small'})[7].getText())
        if yDivYield == 'n/a':
            yDiv:str = 'n/a'
            divYield:str = 'n/a'
        else:
            yDiveYieldList: list[str] = yDivYield.split(' ')
            
            yDiv: float = float(re.search('.*([0-9]\.[0-9]+)',yDiveYieldList[0]).group(1))
            divYield: str = re.search('.*([0-9]\.[0-9]+)',yDiveYieldList[1]).group(1)

        return parsePrice, parsePE, yDiv, divYield

async def getIncome(ticker, session) -> tuple[list[str]]:
    url: str = f'https://stockanalysis.com/stocks/{ticker}/financials/'.format(ticker)
    async with session.get(url) as resp:
        htmlRaw: str = str(await resp.read())
        parse: BeautifulSoup = BeautifulSoup(htmlRaw, 'html.parser')
        
        parseDate: list[str] = list(parse.find_all(name="th", attrs={'class': 'border-b'}))
        parseRev: list[str] = list(parse.find_all(name="td", attrs={'class': 'font-semibold svelte-1eo7czq'}))
        parseNetInc: list[str] = list(parse.find_all(name="td", attrs={'class': 'bolded svelte-1eo7czq'}))
        
        titleList: list[str] = list(parse.find_all(name=lambda tag: tag.name == 'td' and 'gap-x-1' in tag.get("class")))
        cellList: list[str] = list(parse.find_all(name=lambda tag: tag.name == 'td' and not 'gap-x-1' in tag.get("class") and not 'px-2' in tag.get("class")))
        count: int = 0

        for i in titleList:
            if i.getText().strip() == r'EPS (Basic)':
                break
            count += 1
        titleInd: int = count*5

        parseDate = [x.get_text() for x in parseDate if parseDate.index(x) < 6 and parseDate.index(x) > 0]
        revList = [x.get_text() for x in parseRev if parseRev.index(x) < 5]
        netInclist = [x.get_text() for x in parseNetInc if parseNetInc.index(x) < 5]
        epsList = [x.getText() for x in cellList if cellList.index(x) > titleInd - 1 and cellList.index(x)< titleInd + 5]

        return parseDate, revList, netInclist, epsList

async def getBalance(ticker, session) -> tuple[str]:
    url: str = f'https://stockanalysis.com/stocks/{ticker}/financials/balance-sheet/?p=quarterly'.format(ticker)
    async with session.get(url) as resp:
        htmlRaw: str = str(await resp.read())
        parse: BeautifulSoup = BeautifulSoup(htmlRaw, 'html.parser')
        parseQDate: str = str(parse.find(name=lambda tag: tag.name == "th" and tag.get("class") == ['border-b']).getText())
        parseCSI: str = ''
        parseInv: str = ''
        parseCurDebt: str = ''
        parseLongDebt: str = ''
        parseEquity: str = ''

        titleList: list[str] = list(parse.find_all(name=lambda tag: tag.name == 'td' and 'gap-x-1' in tag.get("class")))
        count: int = 0
        cellList: list[str] = list(parse.find_all(name=lambda tag: tag.name == 'td' and not 'gap-x-1' in tag.get("class") and not 'px-2' in tag.get("class")))
        
        for i in titleList:
            findInd: int  = count*20
            match i.getText().strip():
                case r'Cash & Equivalents':
                    parseCSI = str(cellList[findInd].getText())
                case r'Inventory':
                    parseInv = str(cellList[findInd].getText())
                case r'Current Debt':
                    parseCurDebt = str(cellList[findInd].getText())
                case r'Long-Term Debt':
                    parseLongDebt = str(cellList[findInd].getText())
                case r'Shareholders\' Equity':
                    parseEquity = str(cellList[findInd].getText())
            count += 1

        return parseQDate, parseCSI, parseInv, parseCurDebt, parseLongDebt, parseEquity

async def getCash(ticker, session) -> tuple[list[str]]:
    url: str = f'https://stockanalysis.com/stocks/{ticker}/financials/cash-flow-statement/'.format(ticker)
    async with session.get(url) as resp:
        htmlRaw: str = str(await resp.read())
        parse: BeautifulSoup = BeautifulSoup(htmlRaw, 'html.parser')
        parseOpCash: list[str] = []
        parseCapExp: list[str] = []
        parseDiv: list[str] = []
        parseFCF: list[str] = []

        balFinder: list[str] = list(parse.find_all(name=lambda tag: tag.name == 'td' and 'gap-x-1' in tag.get("class")))
        count: int = 0
        cellList: list[str] = list(parse.find_all(name=lambda tag: tag.name == 'td' and not 'gap-x-1' in tag.get("class") and not 'px-2' in tag.get("class")))
        for i in balFinder:
            findInd: int  = count*5
            match i.getText().strip():
                case r'Operating Cash Flow':
                    parseOpCash = [x.getText() for x in cellList if cellList.index(x) > findInd - 1 and cellList.index(x)< findInd + 5]
                case r'Capital Expenditures':
                    parseCapExp = [x.getText() for x in cellList if cellList.index(x) > findInd - 1 and cellList.index(x)< findInd + 5]
                case r'Dividends Paid':
                    parseDiv = [x.getText() for x in cellList if cellList.index(x) > findInd - 1 and cellList.index(x)< findInd + 5]
                case r'Free Cash Flow':
                    parseFCF = [x.getText() for x in cellList if cellList.index(x) > findInd - 1 and cellList.index(x)< findInd + 5]
            count += 1
        
        return parseOpCash, parseCapExp, parseDiv, parseFCF

async def main() -> None:
    ticker: str = sys.argv[1]
    async with aiohttp.ClientSession() as session:
        titleOver: tuple[str] = asyncio.create_task(titleStats(ticker, session))
        incomeStatement: tuple[list[str]] = asyncio.create_task(getIncome(ticker, session))
        balanceSheet: tuple[str] = asyncio.create_task(getBalance(ticker, session))
        cashStatement: tuple[list[str]] = asyncio.create_task(getCash(ticker, session))

        ato = await titleOver
        ais = await incomeStatement 
        abs = await balanceSheet 
        acs = await cashStatement 
     
        workbook: Workbook = Workbook()
        sheet = workbook.active
        colLetterList = ['C', 'D', 'E', 'F', 'G']

        sheet['A1'] = str(sys.argv[1]).upper()
        sheet['B1'] = str(date.today())
        sheet['A2'] = "Price"
        sheet['B2'] = ato[0]
        sheet['A6'] = "DIVIDENDS"
        sheet['A7'] = "Q. Div"
        if ato[2] != 'n/a':
            sheet['B7'] = ato[2]/4.0
        else:
            sheet['B7'] = ato[2]
        
        sheet['A8'] = "Y. Div"
        sheet['B8'] = ato[2]
        sheet['A9'] = "Yield"
        sheet['B9'] = ato[3]

        sheet['E1'] = "PE Ratio"
        sheet['F1'] = ato[1]

        sheet['A11'] = "INCOME STATEMENT (in mln)"
        sheet['A13'] = "Revenue"
        sheet['A14'] = "Net Income"
        sheet['A15'] = "EPS(Basic)"
        countList:int = 0
        rowNum:int = 12
        countRow: int = 0
        for listElem in ais:
            for x in listElem:
                sheet[colLetterList[countList] + str(rowNum + countRow)] = x
                countList += 1
            countRow += 1
            countList = 0

        sheet['A17'] = "BALANCE SHEET (in mln)"
        sheet['B18'] = abs[0]
        sheet['A19'] = "Cash"
        sheet['B19'] = abs[1]
        sheet['A20'] = "Inventory"
        sheet['B20'] = abs[2]
        sheet['A21'] = "Current Portion of LT Debt"
        sheet['B21'] = abs[3]
        sheet['A22'] = "LT Debt"
        sheet['B22'] = abs[4]
        sheet['A23'] = "Equity"  
        sheet['B23'] = abs[5]

        sheet['A25'] = "CASH FLOW (in mln)"
        sheet['A27'] = "Net Operating Cash"
        sheet['A28'] = "Capital Expenditures"
        sheet['A29'] = "Dividends"
        sheet['A30'] = "Free Cash Flow"
        countList2:int = 0
        rowNum:int = 27
        countRow2: int = 0
        for listElem in acs:
            for x in listElem:
                if rowNum == 27:
                    sheet[colLetterList[countList2] + '26'] = ais[0][countList2]
                sheet[colLetterList[countList2] + str(rowNum + countRow2)] = x
                countList2 += 1
            countRow2 += 1
            countList2 = 0

        workbook.save(filename=r"C:\Users\computer1\Desktop\stockoverviewproject\outputs\{}.xlsx".format(ticker + " {}".format(date.today())))

if __name__ == '__main__':
    asyncio.run(main())