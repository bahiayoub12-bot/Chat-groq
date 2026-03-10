# Groq Chat 🚀

واجهة دردشة كاملة مع Groq API تعمل على Docker.

## هيكل المشروع

```
groq-chat/
├── backend/
│   ├── main.py           # FastAPI + Groq
│   └── requirements.txt
├── frontend/
│   └── index.html        # واجهة HTML واحدة
├── Dockerfile
├── docker-compose.yml
└── .env.example
```

## التشغيل

### 1. أضف مفتاح Groq
```bash
cp .env.example .env
# عدّل الملف وضع مفتاحك
```

### 2. بناء وتشغيل
```bash
docker-compose up -d --build
```

### 3. افتح المتصفح
```
http://localhost:8000
```

## أوامر مفيدة

```bash
# عرض اللوجات
docker-compose logs -f

# إيقاف
docker-compose down

# إعادة البناء
docker-compose up -d --build
```

## النماذج المتاحة
- `llama-3.3-70b-versatile` — الأقوى
- `llama-3.1-8b-instant` — الأسرع
- `mixtral-8x7b-32768` — سياق طويل
- `gemma2-9b-it` — خفيف وسريع
