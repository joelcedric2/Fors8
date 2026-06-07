"""Generate a Word document listing every data source the Fors8 prediction system
uses, where it was found, and how it is processed."""

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

OUT = "/Users/joelc/Documents/Github/Fors8/docs/Fors8_Data_Sources.docx"


def add_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    return h


def add_para(doc, text, bold=False, italic=False, size=11):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    return p


def add_source_table(doc, rows, headers=("Source", "Type", "Endpoint / Location", "Auth", "Notes")):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Light Grid Accent 1"
    hdr = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = h
        for run in hdr[i].paragraphs[0].runs:
            run.bold = True
    for row in rows:
        cells = table.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = str(val)
    doc.add_paragraph()


def main():
    doc = Document()

    # Title
    title = doc.add_heading("Fors8 — Data Sources Inventory", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = subtitle.add_run(
        "Geopolitical Conflict Prediction Engine — every external source the "
        "system ingests, where it lives, and how it is used."
    )
    sr.italic = True
    sr.font.size = Pt(11)

    doc.add_paragraph()

    add_heading(doc, "1. Overview", level=1)
    add_para(
        doc,
        "Fors8 grounds its agent-based simulation in real-world data. The pipeline pulls "
        "from four broad categories: (1) open-source intelligence (OSINT) news and event "
        "feeds, (2) financial market data, (3) prediction-market odds for calibration, and "
        "(4) official social-media accounts of governments and military organizations. All "
        "raw records flow through the data ingestor, get tagged with a credibility tier, and "
        "are written into the Zep GraphRAG knowledge graph that the simulation agents read "
        "from at inference time.",
    )

    # ------------------------------------------------------------------
    # SECTION 2 — OSINT NEWS & EVENT FEEDS
    # ------------------------------------------------------------------
    add_heading(doc, "2. OSINT News & Event Feeds", level=1)
    add_para(
        doc,
        "Defined in backend/app/services/osint_scrapers.py. Each scraper is a Python class "
        "wrapped by the OSINTManager which orchestrates fetch cycles, deduplication, and "
        "ingestion into the knowledge graph.",
        italic=True,
    )

    add_heading(doc, "2.1 GDELT — Global Database of Events, Language, and Tone", level=2)
    add_source_table(
        doc,
        [
            (
                "GDELT 2.0",
                "Event database",
                "https://api.gdeltproject.org/api/v2",
                "None (public)",
                "Conflict events, diplomatic meetings, news in 100+ languages",
            ),
        ],
    )
    add_para(
        doc,
        "Implementation: GDELTScraper class. Provides structured event records used as "
        "ground-truth conflict signals fed into agent reasoning.",
    )

    add_heading(doc, "2.2 News RSS Feeds (always-free, no API key)", level=2)
    add_source_table(
        doc,
        [
            ("Al Jazeera", "RSS", "https://www.aljazeera.com/xml/rss/all.xml", "None", "Tier-1 news"),
            ("Reuters World", "RSS", "https://www.reutersagency.com/feed/?taxonomy=best-sectors&post_type=best", "None", "Tier-1 news"),
            ("AP News", "RSS", "https://rsshub.app/apnews/topics/world-news", "None", "Tier-1 news (via rsshub)"),
            ("Al Jazeera Middle East", "RSS", "https://www.aljazeera.com/xml/rss/all.xml", "None", "Regional"),
            ("Middle East Eye", "RSS", "https://www.middleeasteye.net/rss", "None", "Regional"),
            ("The National (UAE)", "RSS", "https://www.thenationalnews.com/rss", "None", "Regional / GCC"),
            ("BBC World", "RSS", "http://feeds.bbci.co.uk/news/world/rss.xml", "None", "Tier-2"),
            ("CNN World", "RSS", "http://rss.cnn.com/rss/edition_world.rss", "None", "Tier-2"),
            ("Guardian World", "RSS", "https://www.theguardian.com/world/rss", "None", "Tier-2"),
            ("Foreign Affairs", "RSS", "https://www.foreignaffairs.com/rss.xml", "None", "Institutional analysis"),
            ("War on the Rocks", "RSS", "https://warontherocks.com/feed/", "None", "Institutional analysis"),
            ("CSIS", "RSS", "https://www.csis.org/analysis/feed", "None", "Institutional analysis"),
            ("CNBC", "RSS", "https://www.cnbc.com/id/100727362/device/rss/rss.html", "None", "Economic"),
            ("Bloomberg Markets", "RSS", "https://feeds.bloomberg.com/markets/news.rss", "None", "Economic"),
            ("Defense One", "RSS", "https://www.defenseone.com/rss/", "None", "Military / defense"),
            ("Breaking Defense", "RSS", "https://breakingdefense.com/feed/", "None", "Military / defense"),
        ],
    )
    add_para(
        doc,
        "Implementation: NewsWebsiteScraper.DEFAULT_FEEDS. Articles are filtered by "
        "geopolitical keywords (iran, israel, missile, hormuz, hezbollah, houthi, "
        "ceasefire, nuclear, escalat).",
    )

    add_heading(doc, "2.3 Google News RSS (topic-targeted, no key)", level=2)
    add_source_table(
        doc,
        [
            ("Iran war", "Google News RSS", "news.google.com/rss/search?q=Iran+war+2026", "None", ""),
            ("Iran-Israel strikes", "Google News RSS", "news.google.com/rss/search?q=Iran+Israel+strikes", "None", ""),
            ("Strait of Hormuz", "Google News RSS", "news.google.com/rss/search?q=Strait+of+Hormuz", "None", ""),
            ("Hezbollah", "Google News RSS", "news.google.com/rss/search?q=Hezbollah+2026", "None", ""),
            ("Houthis / Red Sea", "Google News RSS", "news.google.com/rss/search?q=Houthis+Red+Sea", "None", ""),
            ("Oil price war", "Google News RSS", "news.google.com/rss/search?q=oil+price+Iran+war", "None", ""),
            ("Iran nuclear", "Google News RSS", "news.google.com/rss/search?q=Iran+nuclear+program", "None", ""),
            ("Middle East topic", "Google News RSS", "news.google.com/rss/topics/...Middle East...", "None", ""),
        ],
    )
    add_para(doc, "Implementation: GoogleNewsScraper.TOPIC_FEEDS.")

    add_heading(doc, "2.4 Wikipedia Current Events", level=2)
    add_source_table(
        doc,
        [
            (
                "Wikipedia",
                "Daily events portal",
                "https://en.wikipedia.org/api/rest_v1/page/html/Portal:Current_events",
                "None",
                "Community-curated daily summaries",
            ),
        ],
    )
    add_para(doc, "Implementation: WikipediaCurrentEvents class.")

    add_heading(doc, "2.5 Free-with-Key News APIs", level=2)
    add_source_table(
        doc,
        [
            ("NewsAPI.org", "News API", "https://newsapi.org/v2/everything", "API key", "100 req/day free tier"),
            ("NewsData.io", "News API", "https://newsdata.io/api/1/latest", "API key", "200 credits/day, 30+ languages"),
            ("Mediastack", "News API", "http://api.mediastack.com/v1/news", "Access key", "100 req/month, 7,500+ sources"),
            ("GNews.io", "News API", "https://gnews.io/api/v4/search", "Token", "100 req/day, 60,000+ sources"),
        ],
    )

    add_heading(doc, "2.6 ACLED — Armed Conflict Location & Event Data", level=2)
    add_source_table(
        doc,
        [
            (
                "ACLED",
                "Conflict event database",
                "https://api.acleddata.com/acled/read",
                "Email + API key",
                "Battles, protests, violence-against-civilians with lat/lon and fatalities",
            ),
        ],
    )
    add_para(doc, "Implementation: ACLEDScraper. Adds geocoded conflict events to the graph.")

    add_heading(doc, "2.7 YouTube News Channels (API + RSS fallback)", level=2)
    add_source_table(
        doc,
        [
            ("YouTube Data API v3", "Video metadata", "https://www.googleapis.com/youtube/v3/search", "API key", "Primary path"),
            ("YouTube RSS", "Channel feeds", "https://www.youtube.com/feeds/videos.xml?channel_id=…", "None", "Fallback when no API key"),
        ],
    )
    add_para(doc, "Monitored channels (channel_id, name, credibility):")
    add_source_table(
        doc,
        [
            ("UCMy4o6_qFR0tMr-5FYCr8Gw", "YouTube channel", "Breaking Points", "—", "Independent / less establishment"),
            ("UCNye-wNBqNL5ZzHSJj3l8Bg", "YouTube channel", "Al Jazeera English", "—", "Tier-1 (per user pref)"),
            ("UC16niRr50-MSBwiO3YDb3RA", "YouTube channel", "BBC News", "—", "UK perspective"),
            ("UCupvZG-5ko_eiXAupbDfxWw", "YouTube channel", "CNN", "—", "US establishment"),
            ("UCeY0bbntWzzVIaj2z3QigXg", "YouTube channel", "NBC News", "—", "US mainstream"),
            ("UCwnKziETDbHJtx78nIkfYug", "YouTube channel", "Sky News", "—", "UK perspective"),
            ("UCIALMKvObZNtJ68-rmLjvSA", "YouTube channel", "TLDR News Global", "—", "Analysis-focused, neutral"),
            ("UCBi2mrWuNuyYy4gbM6fU18Q", "YouTube channel", "CNBC", "—", "Economic / markets"),
            ("UCHqC-yWZ1kri4YzwRSt6IGA", "YouTube channel", "TRT World", "—", "Turkish perspective"),
            ("UCQ2sg7vS7JkV_15HD-UNRMw", "YouTube channel", "India Today", "—", "Indian perspective"),
        ],
    )
    add_para(
        doc,
        "Implementation: YouTubeScraper. Pulls video titles, descriptions, and "
        "(where available) transcripts via youtube-transcript-api.",
    )

    add_heading(doc, "2.8 Commodity Prices Snapshot", level=2)
    add_source_table(
        doc,
        [
            ("API Ninjas", "Commodity API", "https://api.api-ninjas.com/v1/commodityprice?name=crude_oil", "API key (free)", "Spot crude oil price fallback"),
        ],
    )
    add_para(doc, "Implementation: OilPriceScraper.")

    # ------------------------------------------------------------------
    # SECTION 3 — FINANCIAL MARKET DATA
    # ------------------------------------------------------------------
    add_heading(doc, "3. Financial Market Data (yfinance)", level=1)
    add_para(
        doc,
        "Defined in backend/app/services/market_data.py. Uses the yfinance Python library "
        "to query Yahoo Finance for current price, 1-week and 1-month percent changes. "
        "Cached for 5 minutes. Tickers grouped by signal category.",
        italic=True,
    )

    add_heading(doc, "3.1 Oil & Energy", level=2)
    add_source_table(
        doc,
        [
            ("Brent crude futures", "Futures", "BZ=F (Yahoo Finance)", "None", "Global oil benchmark"),
            ("WTI crude futures", "Futures", "CL=F (Yahoo Finance)", "None", "US oil benchmark"),
        ],
        headers=("Instrument", "Type", "Ticker / Endpoint", "Auth", "Why it matters"),
    )

    add_heading(doc, "3.2 Defense Contractors", level=2)
    add_source_table(
        doc,
        [
            ("Lockheed Martin", "Equity", "LMT", "None", "Defense procurement signal"),
            ("Raytheon Technologies", "Equity", "RTX", "None", "Missile / defense systems"),
            ("Northrop Grumman", "Equity", "NOC", "None", "Stealth / strategic platforms"),
            ("General Dynamics", "Equity", "GD", "None", "Land / naval"),
            ("Boeing", "Equity", "BA", "None", "Aerospace / defense"),
        ],
        headers=("Company", "Type", "Ticker", "Auth", "Why it matters"),
    )

    add_heading(doc, "3.3 GCC Markets", level=2)
    add_source_table(
        doc,
        [
            ("Saudi Tadawul Index", "Index", "^TASI", "None", "Saudi market"),
            ("Dubai DFM", "Index", "DFMGI.AE", "None", "Dubai market"),
            ("Abu Dhabi ADX", "Index", "ADI.AE", "None", "UAE market"),
            ("Qatar QSE", "Index", "GNRI.QA", "None", "Qatar market"),
            ("Kuwait BK", "Index", "BK.KW", "None", "Kuwait market"),
            ("Saudi Aramco", "Equity", "2222.SR", "None", "Largest oil company"),
            ("ADNOC Distribution", "Equity", "ADNOCDIST.AE", "None", "UAE oil distribution"),
            ("Emirates NBD", "Equity", "ENBD.AE", "None", "Major UAE bank"),
        ],
        headers=("Instrument", "Type", "Ticker", "Auth", "Why it matters"),
    )

    add_heading(doc, "3.4 Shipping & Trade", level=2)
    add_source_table(
        doc,
        [
            ("Breakwave Dry Bulk Shipping ETF", "ETF", "BDRY", "None", "Baltic Dry Index proxy"),
            ("SFL Corporation", "Equity", "SFL", "None", "Tanker shipping"),
        ],
        headers=("Instrument", "Type", "Ticker", "Auth", "Why it matters"),
    )

    add_heading(doc, "3.5 Safe-Haven Indicators", level=2)
    add_source_table(
        doc,
        [
            ("Gold futures", "Futures", "GC=F", "None", "Crisis hedge"),
            ("US Dollar Index", "Index", "DX-Y.NYB", "None", "Reserve currency demand"),
            ("VIX", "Index", "^VIX", "None", "Equity volatility / fear"),
            ("10Y Treasury yield", "Yield", "^TNX", "None", "Flight-to-safety signal"),
        ],
        headers=("Instrument", "Type", "Ticker", "Auth", "Why it matters"),
    )

    # ------------------------------------------------------------------
    # SECTION 4 — PREDICTION MARKETS
    # ------------------------------------------------------------------
    add_heading(doc, "4. Prediction Market Odds (Polymarket)", level=1)
    add_para(
        doc,
        "Defined in backend/app/services/polymarket_client.py. Polymarket odds serve as a "
        "calibration baseline. The simulation's predictions are compared against market "
        "consensus; sustained correct divergences are the signal of value.",
        italic=True,
    )
    add_source_table(
        doc,
        [
            ("Polymarket CLOB API", "Order book / market data", "https://clob.polymarket.com", "None (read-only)", "Real-time event probabilities"),
            ("Polymarket Gamma API", "Market metadata", "https://gamma-api.polymarket.com", "None (read-only)", "Market discovery / search"),
            ("Polymarket Docs", "Reference", "https://docs.polymarket.com/", "—", "API documentation"),
        ],
    )

    # ------------------------------------------------------------------
    # SECTION 5 — OFFICIAL SOCIAL MEDIA ACCOUNTS
    # ------------------------------------------------------------------
    add_heading(doc, "5. Official Social Media Accounts", level=1)
    add_para(
        doc,
        "Defined in backend/app/services/social_media_scraper.py. Scrapes statements from "
        "governments and militaries directly. Posts are tagged with the official-credibility "
        "tier in the data ingestor.",
        italic=True,
    )

    add_heading(doc, "5.1 Platforms & Endpoints", level=2)
    add_source_table(
        doc,
        [
            ("Twitter / X API v2", "Social API", "https://api.twitter.com/2/users/by/username/{handle} & .../tweets", "Bearer token"),
            ("Telegram (Telethon)", "Messaging API", "Telegram MTProto via Telethon library", "API ID + hash"),
            ("Government / state RSS", "RSS", "e.g. http://en.kremlin.ru/events/president/news.rss", "None"),
            ("Truth Social", "Social (custom scrape)", "No public API — custom scraping noted", "—"),
        ],
        headers=("Platform", "Type", "Endpoint", "Auth"),
    )

    add_heading(doc, "5.2 Pre-configured Official Accounts", level=2)
    add_source_table(
        doc,
        [
            ("USA / POTUS", "Twitter", "@POTUS"),
            ("USA / SecDef", "Twitter", "@SecDef"),
            ("USA / CENTCOM", "Twitter", "@CENTCOM"),
            ("USA / Trump", "Truth Social", "@realDonaldTrump"),
            ("Israel / PM Netanyahu", "Twitter", "@netanyahu"),
            ("Israel / IDF Spokesperson", "Twitter", "@IDF"),
            ("Israel / MFA", "Twitter", "@IsraelMFA"),
            ("Iran / Khamenei Office", "Twitter", "@khaboronline"),
            ("Iran / MFA", "Twitter", "@ABORONLINE"),
            ("Iran / Tasnim News", "Telegram", "tasaboronline"),
            ("Iran / IRGC Media", "Telegram", "seaboronline"),
            ("Russia / MFA", "Twitter", "@maboronline"),
            ("Russia / Kremlin", "RSS", "http://en.kremlin.ru/events/president/news.rss"),
            ("China / MFA Spokesperson", "Twitter", "@MFA_China"),
            ("Hezbollah / Al-Manar TV", "Telegram", "almanaboronline"),
            ("Houthis / Military channel", "Telegram", "military_houthi"),
        ],
        headers=("Actor", "Platform", "Account ID"),
    )

    # ------------------------------------------------------------------
    # SECTION 6 — INGESTION & CREDIBILITY
    # ------------------------------------------------------------------
    add_heading(doc, "6. Ingestion Pipeline & Credibility Tiering", level=1)
    add_para(
        doc,
        "Defined in backend/app/services/data_ingestor.py. Every record (news article, "
        "social-media post, market quote, conflict event) is normalized into an "
        "IngestedRecord, tagged with source credibility, and pushed into the Zep GraphRAG "
        "knowledge graph used by the simulation agents.",
    )
    add_source_table(
        doc,
        [
            ("OFFICIAL", "Government / military first-party (POTUS, IDF, MFAs, Kremlin RSS)"),
            ("NEWS_TIER1", "Al Jazeera, Reuters, AP, AFP, BBC Arabic, Al Arabiya"),
            ("NEWS_TIER2", "CNN, BBC, NYT, Washington Post, Guardian, Times of Israel, Haaretz"),
            ("SEMI_OFFICIAL", "Iranian state-aligned (Press TV, Tasnim, Fars), IDF/CENTCOM communiqués"),
            ("INSTITUTIONAL", "CSIS, Foreign Affairs, War on the Rocks, Defense One, Breaking Defense, Wikipedia, ACLED"),
            ("OSINT", "Open-source intelligence and satellite imagery"),
            ("UNVERIFIED", "Anything not matching the above"),
        ],
        headers=("Tier", "Sources mapped to it"),
    )

    # ------------------------------------------------------------------
    # SECTION 7 — KNOWLEDGE GRAPH & DOWNSTREAM USE
    # ------------------------------------------------------------------
    add_heading(doc, "7. Knowledge Graph & Downstream Use", level=1)
    add_source_table(
        doc,
        [
            ("Zep Cloud (GraphRAG)", "Knowledge graph", "Zep API (env: ZEP_API_KEY)", "API key", "Stores 612+ entities and 1,200+ relationships"),
            ("Ollama (vLLM)", "Inference", "env: VLLM_ENDPOINT / VLLM_MODEL", "API key", "Hosts qwen2.5:32b on Vast.ai 2x A100 GPUs"),
            ("PostgreSQL", "Persistence", "Local DB", "—", "Simulation runs, predictions, Brier scores"),
        ],
        headers=("System", "Role", "Endpoint / Config", "Auth", "Purpose"),
    )
    add_para(
        doc,
        "Each simulation run reads from the knowledge graph, executes 50,000+ agents across "
        "18 countries (15 specialized roles), aggregates outcomes via Monte Carlo, and writes "
        "calibrated probability-weighted predictions and per-actor outcomes back to PostgreSQL. "
        "Brier scores against resolved predictions track calibration over time "
        "(brier_tracker.py).",
    )

    # ------------------------------------------------------------------
    # SECTION 8 — FILE REFERENCE
    # ------------------------------------------------------------------
    add_heading(doc, "8. Source-Code Reference", level=1)
    add_source_table(
        doc,
        [
            ("backend/app/services/osint_scrapers.py", "GDELT, RSS, Google News, Wikipedia, NewsAPI, NewsData, Mediastack, GNews, ACLED, YouTube, oil price"),
            ("backend/app/services/market_data.py", "yfinance — oil, defense stocks, GCC markets, shipping, safe-havens"),
            ("backend/app/services/polymarket_client.py", "Polymarket CLOB + Gamma APIs"),
            ("backend/app/services/social_media_scraper.py", "Twitter v2, Telegram, government RSS, Truth Social"),
            ("backend/app/services/data_ingestor.py", "Normalization + credibility tiering + Zep ingestion"),
            ("backend/app/services/graph_builder.py", "Knowledge graph construction"),
            ("backend/app/services/prediction_engine.py", "Simulation prediction logic"),
            ("backend/app/services/brier_tracker.py", "Calibration tracking"),
        ],
        headers=("File", "What it provides"),
    )

    doc.save(OUT)
    print("Wrote:", OUT)


if __name__ == "__main__":
    main()
