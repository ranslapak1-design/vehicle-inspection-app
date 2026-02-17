# Vehicle Inspection App

אפליקציית Flask לניהול בדיקות רכב - צילום, מילוי נתונים, ומעקב חוסרים.

## תכונות

- **ניהול יצרנים ורכבים** - ניווט היררכי: יצרן > תאריך > רכב
- **צילום תמונות** - צילום ישיר מהטלפון ושמירה בתיקיית הרכב
- **מילוי נתוני בדיקה** - קריאת נתוני מזכירה מהאקסל והשוואה לנתוני הבוחן
- **חוסרים** - הצגת פערים מגיליון "פ. ממצאים מסכם", הערות בוחן עם שמירה אוטומטית
- **הפקת PDF** - יצירת קובץ PDF עם סיכום חוסרים, נשמר בתיקיית הרכב
- **שיתוף WhatsApp** - שליחת הודעת טקסט מפורטת עם כל החוסרים

## דרישות

- Python 3.10+
- Flask
- openpyxl
- fpdf2
- Pillow

## התקנה

```bash
pip install flask openpyxl fpdf2 pillow
```

## הפעלה

```bash
python app.py
```

השרת יעלה בכתובת `http://0.0.0.0:5555` ונגיש ברשת המקומית.

## מבנה הפרויקט

```
├── app.py                  # שרת Flask ראשי
├── templates/
│   ├── filelist.html        # רשימת יצרנים
│   ├── dates.html           # רשימת תאריכים
│   ├── vehicles.html        # רשימת רכבים
│   ├── category.html        # בחירת קטגוריה
│   ├── inspect.html         # דף בדיקה ראשי (N2/N3)
│   ├── inspect_empty.html   # placeholder לקטגוריות נוספות
│   └── inspect_m2m3.html    # טופס M2/M3
├── inspect_sheets.py        # סקריפטים לניתוח גיליונות
└── inspect_sheets2.py
```

## מבנה תיקיות נתונים

```
יצרנים/
├── <יצרן>/
│   └── <תאריך>/
│       └── <רכב>/
│           ├── <רכב>.xlsx    # קובץ אקסל עם נתוני הבדיקה
│           ├── תמונות/       # תיקיית תמונות
│           ├── <רכב> - חוסרים.pdf   # PDF חוסרים (נוצר אוטומטית)
│           └── <רכב> - חוסרים.png   # תמונת חוסרים (נוצר אוטומטית)
```

## API Endpoints

| Endpoint | Method | תיאור |
|----------|--------|-------|
| `/api/secretary` | GET | נתוני מזכירה מהאקסל |
| `/api/save` | POST | שמירת נתוני בוחן לאקסל |
| `/api/save_photo` | POST | שמירת תמונה |
| `/api/deficiencies` | GET | נתוני חוסרים + הערות בוחן |
| `/api/save_deficiency_notes` | POST | שמירת הערות בוחן |
| `/api/deficiency_pdf` | GET | הפקת PDF חוסרים |
| `/api/deficiency_image` | GET | הפקת תמונת חוסרים |
| `/api/deficiency_text` | GET | טקסט חוסרים לשיתוף WhatsApp |
| `/api/classifications` | GET | אפשרויות סיווג |
| `/api/save_classification` | POST | שמירת סיווג |
