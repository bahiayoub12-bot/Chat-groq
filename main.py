from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel
from groq import Groq
import os, io, base64, tempfile, asyncio, urllib.parse

# مكتبات الملفات
import pypdf
import pdfplumber
from docx import Document
import openpyxl
import pandas as pd
import numpy as np
from PIL import Image

# مكتبات البحث والإنترنت
from duckduckgo_search import DDGS
from youtube_transcript_api import YouTubeTranscriptApi
import httpx
from bs4 import BeautifulSoup

# TTS
try:
    import edge_tts
    EDGE_TTS_AVAILABLE = True
except:
    EDGE_TTS_AVAILABLE = False

from gtts import gTTS

try:
    import pytesseract
    OCR_AVAILABLE = True
except:
    OCR_AVAILABLE = False

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

client = Groq(api_key=os.environ.get("GROQ_API_KEY"))

# ══ System Prompt ══
SYSTEM_PROMPT = """أنت Groq Chat، مساعد ذكاء اصطناعي متطور ومتعدد القدرات.

قدراتك الفعلية:
- تحليل الملفات: PDF, Word, Excel, CSV, الصور
- البحث في الإنترنت: عندما يطلب المستخدم البحث أو الأخبار
- استخراج نصوص فيديوهات يوتيوب: عندما يرسل المستخدم رابط يوتيوب
- قراءة محتوى المواقع: عندما يرسل المستخدم رابط موقع
- التحدث والاستماع بالصوت

قواعد مهمة:
- لا تقل أبداً "لا أستطيع الوصول للإنترنت" لأنك تستطيع ذلك عبر الأدوات المتاحة
- عندما يرسل المستخدم رابط يوتيوب، النص سيُستخرج تلقائياً وسيصلك في السياق
- عندما يطلب البحث، النتائج ستصلك تلقائياً في السياق
- أجب دائماً بالعربية ما لم يطلب المستخدم غير ذلك
- كن مختصراً وواضحاً في إجاباتك
- أنت تعمل على Railway"""

# ══ استخراج النص من الملفات ══

def extract_pdf(data: bytes) -> str:
    text = ""
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        text += " | ".join([str(c) for c in row if c]) + "\n"
    except:
        try:
            reader = pypdf.PdfReader(io.BytesIO(data))
            for page in reader.pages:
                text += page.extract_text() or ""
        except:
            pass
    return text[:8000] if text else "[لم يتم استخراج نص من PDF]"

def extract_docx(data: bytes) -> str:
    doc = Document(io.BytesIO(data))
    text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    for table in doc.tables:
        for row in table.rows:
            text += " | ".join([c.text.strip() for c in row.cells]) + "\n"
    return text[:8000]

def extract_excel(data: bytes) -> str:
    try:
        df_dict = pd.read_excel(io.BytesIO(data), sheet_name=None)
        text = ""
        for sheet_name, df in df_dict.items():
            text += f"\n[ورقة: {sheet_name}]\n"
            text += f"الأبعاد: {df.shape[0]} صف × {df.shape[1]} عمود\n"
            text += f"الأعمدة: {', '.join([str(c) for c in df.columns])}\n"
            text += df.to_string(max_rows=50) + "\n"
            numeric_cols = df.select_dtypes(include=[np.number])
            if not numeric_cols.empty:
                text += "\nإحصائيات:\n" + numeric_cols.describe().to_string() + "\n"
        return text[:8000]
    except:
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True)
        text = ""
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            text += f"\n[ورقة: {sheet}]\n"
            for row in ws.iter_rows(values_only=True):
                vals = [str(v) for v in row if v is not None]
                if vals:
                    text += " | ".join(vals) + "\n"
        return text[:8000]

def extract_csv(data: bytes) -> str:
    try:
        df = pd.read_csv(io.BytesIO(data))
        text = f"عدد الصفوف: {len(df)}\n"
        text += f"الأعمدة: {', '.join(df.columns)}\n"
        text += df.to_string(max_rows=50) + "\n"
        numeric_cols = df.select_dtypes(include=[np.number])
        if not numeric_cols.empty:
            text += "\nإحصائيات:\n" + numeric_cols.describe().to_string()
        return text[:8000]
    except:
        return data.decode("utf-8", errors="ignore")[:8000]

def extract_image(data: bytes, filename: str) -> str:
    if OCR_AVAILABLE:
        try:
            img = Image.open(io.BytesIO(data))
            text = pytesseract.image_to_string(img, lang='ara+eng')
            if text.strip():
                return f"[نص مستخرج من الصورة عبر OCR]\n{text}"
        except:
            pass
    return "__IMAGE__"

def extract_text(filename: str, data: bytes) -> str:
    ext = filename.lower().split(".")[-1]
    if ext == "pdf":                         return extract_pdf(data)
    elif ext == "docx":                      return extract_docx(data)
    elif ext in ("xlsx", "xls"):             return extract_excel(data)
    elif ext == "csv":                       return extract_csv(data)
    elif ext in ("jpg","jpeg","png","webp"): return extract_image(data, filename)
    elif ext in ("txt", "md"):               return data.decode("utf-8", errors="ignore")[:8000]
    return ""

# ══ استخراج نص من موقع ══
async def scrape_url(url: str) -> str:
    try:
        async with httpx.AsyncClient(timeout=10, follow_redirects=True) as c:
            headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
            resp = await c.get(url, headers=headers)
            soup = BeautifulSoup(resp.text, "html.parser")
            for tag in soup(["script","style","nav","footer","header","aside"]):
                tag.decompose()
            text = soup.get_text(separator="\n", strip=True)
            lines = [l.strip() for l in text.splitlines() if len(l.strip()) > 30]
            return "\n".join(lines)[:6000]
    except Exception as e:
        return f"[خطأ في قراءة الموقع: {str(e)}]"

# ══ Models ══
class ChatRequest(BaseModel):
    messages: list
    model: str = "llama-3.3-70b-versatile"

class SearchRequest(BaseModel):
    query: str
    max_results: int = 5

class TTSRequest(BaseModel):
    text: str
    lang: str = "ar"
    speed: float = 1.0

class YouTubeRequest(BaseModel):
    url: str

class ScrapeRequest(BaseModel):
    url: str

# ══ Endpoints ══

@app.post("/api/chat")
async def chat(req: ChatRequest):
    messages = list(req.messages)
    if not messages or messages[0].get("role") != "system":
        messages = [{"role": "system", "content": SYSTEM_PROMPT}] + messages
    response = client.chat.completions.create(
        model=req.model,
        messages=messages,
        max_tokens=1500,
    )
    return {"reply": response.choices[0].message.content}

@app.post("/api/upload")
async def upload(file: UploadFile = File(...)):
    data = await file.read()
    text = extract_text(file.filename, data)
    if text == "__IMAGE__":
        b64 = base64.b64encode(data).decode()
        ext = file.filename.lower().split(".")[-1]
        mime = "image/jpeg" if ext in ("jpg","jpeg") else f"image/{ext}"
        response = client.chat.completions.create(
            model="llama-3.2-11b-vision-preview",
            messages=[{"role":"user","content":[
                {"type":"image_url","image_url":{"url":f"data:{mime};base64,{b64}"}},
                {"type":"text","text":"صف محتوى هذه الصورة بالتفصيل بالعربية"}
            ]}],
            max_tokens=1024
        )
        return {"text": response.choices[0].message.content, "filename": file.filename, "type": "image"}
    if not text:
        return {"text": "", "filename": file.filename, "message": "لم يتم استخراج نص"}
    return {"text": text, "filename": file.filename, "type": "document"}

@app.post("/api/search")
async def search(req: SearchRequest):
    try:
        results = []
        with DDGS() as ddgs:
            for r in ddgs.text(req.query, max_results=req.max_results):
                results.append(f"• {r['title']}\n{r['body']}\n{r['href']}")
        return {"results": "\n\n".join(results), "count": len(results)}
    except Exception as e:
        return {"results": "", "error": str(e)}

@app.post("/api/youtube")
async def youtube(req: YouTubeRequest):
    try:
        url = req.url.strip()
        video_id = ""

        if "youtu.be/" in url:
            video_id = url.split("youtu.be/")[1].split("?")[0].split("/")[0]
        elif "youtube.com/watch" in url:
            params = urllib.parse.parse_qs(urllib.parse.urlparse(url).query)
            video_id = params.get("v", [""])[0]
        elif "youtube.com/shorts/" in url:
            video_id = url.split("youtube.com/shorts/")[1].split("?")[0]

        video_id = video_id.strip()
        if not video_id:
            return {"transcript": "", "error": "لم يتم التعرف على معرف الفيديو"}

        # محاولة استخراج بعدة طرق
        transcript_list = None
        for langs in [["ar"], ["en"], ["ar", "en"], ["fr"], ["auto"]]:
            try:
                if langs == ["auto"]:
                    available = YouTubeTranscriptApi.list_transcripts(video_id)
                    for t in available:
                        transcript_list = t.fetch()
                        break
                else:
                    transcript_list = YouTubeTranscriptApi.get_transcript(video_id, languages=langs)
                if transcript_list:
                    break
            except:
                continue

        if not transcript_list:
            return {"transcript": "", "error": f"لا تتوفر ترجمة لهذا الفيديو (ID: {video_id})"}

        transcript = " ".join([t["text"] for t in transcript_list])
        return {"transcript": transcript[:8000], "video_id": video_id}

    except Exception as e:
        return {"transcript": "", "error": str(e)}

@app.post("/api/scrape")
async def scrape(req: ScrapeRequest):
    content = await scrape_url(req.url)
    return {"content": content, "url": req.url}

@app.post("/api/tts")
async def tts(req: TTSRequest):
    from fastapi.responses import Response
    text = req.text[:600]
    speed = max(0.5, min(2.0, req.speed))

    if EDGE_TTS_AVAILABLE:
        try:
            rate_percent = int((speed - 1.0) * 100)
            rate_str = f"+{rate_percent}%" if rate_percent >= 0 else f"{rate_percent}%"
            communicate = edge_tts.Communicate(text, "ar-SA-ZariyahNeural", rate=rate_str)
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".mp3")
            await communicate.save(tmp.name)
            with open(tmp.name, "rb") as f:
                audio_bytes = f.read()
            os.unlink(tmp.name)
            return Response(
                content=audio_bytes,
                media_type="audio/mpeg",
                headers={
                    "Access-Control-Allow-Origin": "*",
                    "Access-Control-Allow-Methods": "POST, OPTIONS",
                    "Access-Control-Allow-Headers": "*",
                    "Content-Disposition": "inline; filename=reply.mp3"
                }
            )
        except Exception as e:
            pass

    try:
        tts_obj = gTTS(text=text, lang=req.lang, slow=(speed < 0.8))
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".mp3")
        tts_obj.save(tmp.name)
        with open(tmp.name, "rb") as f:
            audio_bytes = f.read()
        os.unlink(tmp.name)
        return Response(
            content=audio_bytes,
            media_type="audio/mpeg",
            headers={
                "Access-Control-Allow-Origin": "*",
                "Access-Control-Allow-Methods": "POST, OPTIONS",
                "Access-Control-Allow-Headers": "*",
                "Content-Disposition": "inline; filename=reply.mp3"
            }
        )
    except Exception as e:
        return {"error": str(e)}

@app.get("/health")
def health():
    return {"status": "ok", "ocr": OCR_AVAILABLE, "edge_tts": EDGE_TTS_AVAILABLE, "platform": "railway"}

app.mount("/", StaticFiles(directory="/app/frontend", html=True), name="frontend")

@app.get("/test")
async def test_all():
    results = {}

    # 1. edge-tts
    try:
        import edge_tts
        results["edge_tts"] = "✅ متاح"
    except Exception as e:
        results["edge_tts"] = f"❌ {str(e)[:60]}"

    # 2. DuckDuckGo Search
    try:
        from duckduckgo_search import DDGS
        with DDGS() as ddgs:
            r = list(ddgs.text("اختبار", max_results=1))
        results["duckduckgo"] = f"✅ يعمل — {r[0]['title'][:40] if r else 'لا نتائج'}"
    except Exception as e:
        results["duckduckgo"] = f"❌ {str(e)[:80]}"

    # 3. YouTube Transcript
    try:
        from youtube_transcript_api import YouTubeTranscriptApi
        t = YouTubeTranscriptApi.get_transcript("YQHsXMglC9A", languages=["ar","en"])
        results["youtube_transcript"] = f"✅ يعمل — {len(t)} جملة"
    except Exception as e:
        results["youtube_transcript"] = f"❌ {str(e)[:80]}"

    # 4. Scrapy
    try:
        import scrapy
        results["scrapy"] = f"✅ متاح — v{scrapy.__version__}"
    except Exception as e:
        results["scrapy"] = f"❌ {str(e)[:60]}"

    # 5. BeautifulSoup + httpx
    try:
        import httpx
        from bs4 import BeautifulSoup
        async with httpx.AsyncClient(timeout=8) as c:
            resp = await c.get("https://example.com")
        soup = BeautifulSoup(resp.text, "html.parser")
        title = soup.find("title").text if soup.find("title") else "لا عنوان"
        results["beautifulsoup_httpx"] = f"✅ يعمل — {title[:40]}"
    except Exception as e:
        results["beautifulsoup_httpx"] = f"❌ {str(e)[:80]}"

    # 6. pdfplumber
    try:
        import pdfplumber
        results["pdfplumber"] = "✅ متاح"
    except Exception as e:
        results["pdfplumber"] = f"❌ {str(e)[:60]}"

    # 7. pandas + numpy
    try:
        import pandas as pd
        import numpy as np
        results["pandas_numpy"] = f"✅ pandas {pd.__version__} / numpy {np.__version__}"
    except Exception as e:
        results["pandas_numpy"] = f"❌ {str(e)[:60]}"

    # 8. Groq API
    try:
        r = client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[{"role":"user","content":"قل: اختبار ناجح"}],
            max_tokens=10
        )
        results["groq_api"] = f"✅ {r.choices[0].message.content}"
    except Exception as e:
        results["groq_api"] = f"❌ {str(e)[:80]}"

    return {"platform": "railway", "tests": results}
