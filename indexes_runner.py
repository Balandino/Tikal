# pylint: disable=line-too-long

"""Output Market indexes into Excel"""
import time

from indexes_utilities import add_indexes, add_title_page, get_all_historical_data
from workbook_utilities import close_workbook, create_workbook

START_TIME = time.perf_counter()

INDEXES = [
    [
        "^GSPC",
        "S&P 500",
        "U.S",
        "The S&P 500 Index, or Standard & Poor's 500 Index, is a market-capitalization-weighted index of 500 leading publicly traded companies in the U.S.",
        "Index",
    ],
    [
        "^IXIC",
        "Nasdaq Composite",
        "U.S Nasdaq Pre Market Indicator",
        "The Nasdaq 100 Index is a basket of the 100 largest, most actively traded U.S companies listed on the Nasdaq stock exchange",
        "Index",
    ],
    [
        "^NDX",
        "Nasdaq 100",
        "U.S Tech Stocks",
        "The Nasdaq Composite is a market capitalization-weighted index of more than 3,700 stocks listed on the Nasdaq stock exchange",
        "Index",
    ],
    [
        "^DJI",
        "Dow Jones",
        "U.S Blue Chip Stocks",
        "The Dow Jones Industrial Average (DJIA), also known as the Dow 30, is a stock market index that tracks 30 large, publicly-owned blue-chip companies trading on the New York Stock Exchange (NYSE) and Nasdaq",
        "Index",
    ],
    [
        "^RUT",
        "Russell 2000",
        "U.S Small Cap",
        "The Russell 2000 Index is a small-cap U.S. stock market index that makes up the smallest 2,000 stocks in the Russell 3000 Index.",
        "Index",
    ],
    [
        "^RUA",
        "Russell 3000",
        "U.S Stock Market",
        "A capitalization-weighted stock market index that seeks to be a benchmark of the entire U.S. stock market. It measures the performance of the 3,000 largest publicly held companies incorporated in America as measured by total market capitalization",
        "Index",
    ],
    [
        "^STOXX50E",
        "Eurostoxx 50",
        "Europe",
        "The EURO STOXX 50 Index is a market capitalization-weighted stock index of 50 large, blue-chip European companies operating within eurozone nations",
        "Index",
    ],
    [
        "^STOXX",
        "Eurostoxx 600",
        "Europe",
        "The Stoxx Europe 600 Index is derived from STOXX's Europe Total Market Index and is a subset of the popular Stoxx Global 1800 Index.  It has a fixed number of 600 components, representing large, mid, and small-capitalization companies from 17  countries",
        "Index",
    ],
    [
        "^FTSE",
        "FTSE 100",
        "UK",
        "A share index of the 100 companies listed on the London Stock Exchange with (in principle) the highest market capitalisation",
        "Index",
    ],
    [
        "^FTMC",
        "FTSE 250",
        "UK",
        "A capitalisation-weighted index consisting of the 101st to the 350th largest companies listed on the London Stock Exchange",
        "Index",
    ],
    [
        "^FTLC",
        "FTSE 350",
        "UK",
        "A market capitalization weighted stock market index made up of the constituents of the FTSE 100 and FTSE 250 indices",
        "Index",
    ],
    [
        "^GSPTSE",
        "S&P-TSX Composite",
        "Canada",
        "The S&P/TSX Composite is the headline index for the Canadian equity market.",
        "Index",
    ],
    [
        "^GDAXI",
        "DAX 40",
        "Germany",
        "The DAX—also known as the Deutscher Aktien Index or the GER40—is a stock index that represents 40 of the largest and most liquid German companies that trade on the Frankfurt Exchange",
        "Index",
    ],
    [
        "^HSI",
        "HSI",
        "Hong Kong",
        "The Hang Seng Index or HSI is a free-float market capitalization-weighted index of the sixty largest companies that trade on the Hong Kong Exchange (HKEx)",
        "Index",
    ],
    [
        "^HSCE",
        "HSCEI",
        "Hong Kong Exchange",
        "China 50 H-Share stocks listed on Hang Seng China Enterprises Index is a stock market index of The Stock Exchange of Hong Kong for H share, red chip, and P chip",
        "Index",
    ],
    [
        "^N225",
        "Nikkei 225",
        "Japan",
        " It is a price-weighted index composed of Japan's top 225 blue-chip companies traded on the Tokyo Stock Exchange. The Nikkei is equivalent to the Dow Jones Industrial Average (DJIA) Index in the United States",
        "Index",
    ],
    [
        "^AXJO",
        "ASX 200",
        "Australia",
        "The S&P/ASX 200 Index is the benchmark institutional investable stock market index in Australia, comprising the 200 largest stocks by float-adjusted market capitalization",
        "Index",
    ],
    [
        "^VIX",
        "Vix",
        "U.S (The 'Fear' Index)",
        "The Cboe Volatility Index, or VIX, is a real-time market index representing the market’s expectations for volatility over the coming 30 days",
        "Vol",
    ],
    [
        "^VXD",
        "Dow Jones Vol",
        "U.S",
        "The Cboe DJIA Volatility Index is a VIX-style estimate of the expected 30-day volatility of DJIA stock index returns",
        "Vol",
    ],
    [
        "^VOLQ",
        "Nasdaq 100 Vol",
        "U.S Tech",
        "The Nasdaq-100 Volatility Index measures changes in 30 day implied volatility of the Nasdaq-100 index",
        "Vol",
    ],
    [
        "^RVX",
        "Russell 2000 vol",
        "-",
        "The Cboe Russell 2000 Volatility IndexSM (RVX) is a VIX-style estimate of the expected 30-day volatility of Russell 2000 Index returns",
        "Vol",
    ],
    [
        "^OVX",
        "Crude Oil Vol",
        "-",
        "The Cboe Crude Oil ETF Volatility IndexSM is an estimate of the expected 30-day volatility of crude oil as priced by the United States Oil Fund (USO)",
        "Vol",
    ],
    [
        "^GVZ",
        "Gold Volatility",
        "-",
        "The Cboe Gold ETF Volatility IndexSM (GVZ) is an estimate of the expected 30-day volatility of returns on the SPDR Gold Shares ETF (GLD).",
        "Vol",
    ],
    [
        "XLY",
        "Consumer Discretionary",
        "U.S",
        "XLY tracks a market-cap-weighted index of consumer-discretionary stocks drawn from the S&P 500.  Note Amazon & Tesla form roughly 20% each, so influence the index",
        "Etf",
    ],
    [
        "XLP",
        "Consumer Staples",
        "U.S",
        "A Consumer Staples ETF, XLP pulls its stocks from the S&P 500 rather than the broad market.  This produces somewhat-concentrated exposure. The fund's holdings are nearly all large-caps",
        "Etf",
    ],
    [
        "XLE",
        "Energy",
        "U.S",
        "XLE offers liquid exposure to a market-like basket of US energy firms. “Market-like” in the context of the energy sector means concentrated exposure to the giants in the industry",
        "Etf",
    ],
    [
        "XLF",
        "Financials",
        "U.S",
        "XLF tracks an index of S&P 500 financial stocks, weighted by market cap.",
        "Etf",
    ],
    [
        "XLV",
        "Health Care",
        "U.S",
        "XLV tracks health care stocks from within the S&P 500 Index, weighted by market cap",
        "Etf",
    ],
    [
        "XLI",
        "Industrials",
        "U.S",
        "XLI tracks a market cap-weighted index of industrial-sector stocks drawn from the S&P 500.",
        "Etf",
    ],
    [
        "XLB",
        "Materials",
        "U.S",
        "XLB invests in basic materials companies from the S&P 500, such as chemicals, metals and mining, paper and forest products, containers and packaging, and construction materials industry. ",
        "Etf",
    ],
    [
        "XLRE",
        "Real Estate",
        "U.S",
        "XLRE tracks a market-cap-weighted index of REITs and real estate stocks, excluding mortgage REITs, from the S&P 500. ",
        "Etf",
    ],
    [
        "XLK",
        "Technology",
        "U.S",
        "XLK tracks an index of S&P 500 technology stocks.",
        "Etf",
    ],
    [
        "XLC",
        "Communication",
        "U.S",
        "XLC tracks a market-cap-weighted index of US telecommunication and media & entertainment components of the S&P 500 index.",
        "Etf",
    ],
    [
        "XLU",
        "Utilities",
        "U.S",
        "XLU tracks a market-cap-weighted index of US utilities stocks drawn exclusively from the S&P 500.",
        "Etf",
    ],
    [
        "XME",
        "Mining",
        "U.S",
        "XME tracks an equal-weighted index of US metals and mining companies.",
        "Etf",
    ],
    [
        "VNQ",
        "REITS",
        "U.S",
        "VNQ tracks a market-cap-weighted index of companies involved in the ownership and operation of real estate in the United States.",
        "Etf",
    ],
    [
        "GDX",
        "Gold Miners",
        "U.S",
        "GDX tracks a market-cap-weighted index of global gold-mining firms. ",
        "Etf",
    ],
    [
        "AMLP",
        "Energy Infastructure",
        "U.S",
        "AMLP tracks a market-cap-weighted index of publicly-traded energy infrastructure MLPs in the US.",
        "Etf",
    ],
    [
        "ITB",
        "Homebuilders",
        "U.S",
        "ITB tracks a market-cap-weighted index of companies involved in the production and sale of materials used in home construction.",
        "Etf",
    ],
    [
        "OIH",
        "Oil Services",
        "U.S",
        "OIH tracks a market-cap-weighted index of 25 of the largest US-listed, publicly traded oil services companies.",
        "Etf",
    ],
    [
        "KRE",
        "Regional Banks",
        "U.S",
        "KRE tracks an equal-weighted index of US regional banking stocks.",
        "Etf",
    ],
    [
        "XRT",
        "Retail",
        "U.S",
        "XRT tracks a broad-based, equal-weighted index of stocks in the US retail industry.",
        "Etf",
    ],
    [
        "MOO",
        "Agriculture",
        "U.S",
        "MOO tracks a market-cap-weighted index of companies that generate revenues from the agribusiness sector.",
        "Etf",
    ],
    [
        "FDN",
        "Internet",
        "U.S",
        "FDN tracks a market-cap-weighted index of the largest and most liquid US Internet companies.",
        "Etf",
    ],
    [
        "IBB",
        "Biotech",
        "U.S",
        "IBB tracks the performance of a modified market-cap-weighted index of US biotechnology companies listed on US exchanges. ",
        "Etf",
    ],
    [
        "SMH",
        "Semiconductors",
        "U.S",
        "SMH tracks a market-cap-weighted index of 25 of the largest US-listed semiconductors companies.",
        "Etf",
    ],
    [
        "XOP",
        "Oil & Gas E&P",
        "U.S",
        "XOP tracks an equal-weighted index of companies in the US oil & gas exploration & production space.",
        "Etf",
    ],
    [
        "KIE",
        "Insurance",
        "U.S",
        "KIE tracks an equal-weighted-index of insurance companies, as defined by GICS.",
        "Etf",
    ],
    [
        "PHO",
        "Water Resources",
        "U.S",
        "PHO tracks a modified liquidity-weighted index of US-listed companies that create products to conserve and purify water.",
        "Etf",
    ],
    [
        "IGV",
        "Software",
        "U.S/Canada",
        "IGV tracks a market-cap-weighted index of US and Canadian software companies",
        "Etf",
    ],
    [
        "TAN",
        "Solar",
        "Global",
        "TAN tracks an index of global solar energy companies selected based on the revenue generated from solar related business.",
        "Etf",
    ],
    [
        "PBW",
        "Clean Energy",
        "U.S",
        "PBW tracks a modified equal-weighted index of companies involved in cleaner energy sources or energy conservation.",
        "Etf",
    ],
    [
        "JETS",
        "Airlines",
        "Global",
        "JETS invests in both US and non-US airline companies. This concentrated portfolio is weighted towards domestic passenger airlines.",
        "Etf",
    ],
    [
        "ITA",
        "Aerospace and Defense",
        "U.S",
        "Companies in this sector tend to be rather large, slow growing, but remarkably stable due to the widespread use of long-term government contracts for most of their services.",
        "Etf",
    ],
    [
        "IAU",
        "iShares Gold Trust",
        "Gold",
        "The iShares Gold Trust (the 'Trust') seeks to reflect generally the performance of the price of gold",
        "Etf",
    ],
    [
        "IWM",
        "iShares Russell 2000 ETF",
        "U.S",
        "The iShares Russell 2000 ETF seeks to track the investment results of an index composed of small-capitalization U.S. equities.",
        "Etf",
    ],
    [
        "URA",
        "Global X Uranium",
        "Uranium",
        "The Fund seeks to provide investment results that correspond generally to the price and yield performance of the Solactive Global Uranium Index.",
        "Etf",
    ],
    # [
    #     "WTI",
    #     "Crude Oil",
    #     "-",
    #     "West Texas Intermediate",
    #     "Commodity",
    # ],
    # [
    #     "GC=F",
    #     "Gold",
    #     "-",
    #     "Gold",
    #     "Commodity",
    # ],
]


COLORS = {
    "Index": "#ff4f4f",
    "Vol": "#fff2cc",
    "Etf": "#50e0ff",
    "Commodity": "#ffbf00",
}

WORKBOOK_NAME = "Workbooks/Indexes.xlsx"
WORKBOOK = create_workbook(WORKBOOK_NAME)


tickers: list[str] = []
for index_data in INDEXES:
    tickers.append(index_data[0])

# Obtain all historical data at once to avoid sequential slowdown
print("[Downloading] Obtaining historical data")
HISTORICAL = get_all_historical_data(tickers)
print("[Downloading] Data obtained\n")


# Add header page
header_page = add_title_page(WORKBOOK, INDEXES, COLORS, HISTORICAL)

add_indexes(WORKBOOK, INDEXES, HISTORICAL, COLORS)

# Save and close
print("\n[Writing] Writing Workbook")
close_workbook(WORKBOOK, WORKBOOK_NAME)
print("[Writing] Workbook written!")
print("\n[System] Opening Workbook")
print(f"\n[TIMING] Total time elapsed: {time.perf_counter()-START_TIME:0.2f} seconds")

