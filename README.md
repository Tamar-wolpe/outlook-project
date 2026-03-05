# 📧 Outlook Draft Generator

מערכת חכמה המאפשרת יצירת טיוטות ב-Microsoft Outlook ישירות דרך ממשק דפדפן (Web). המערכת תומכת בשליחה למספר נמענים ובצירוף קבצים באופן אוטומטי.



## ✨ תכונות עיקריות
* **ממשק משתמש מודרני:** עיצוב נקי, תמיכה בעברית וחווית משתמש נוחה (RTL).
* **תמיכה בקבצים:** העלאת קבצי PDF, Word, Excel ותמונות (עד 10MB).
* **נמענים מרובים:** אפשרות להזנת רשימת מיילים מופרדת בפסיקים.
* **אוטומציה מלאה:** שימוש ב-VBScript לחיבור ישיר לאפליקציית Outlook המותקנת במחשב.

## 🛠 טכנולוגיות
* **Frontend:** HTML5, CSS3 (Flexbox/Grid), JavaScript (Vanilla).
* **Backend:** Node.js, Express.
* **Automation:** VBScript / PowerShell integration.

---

## 🚀 הוראות התקנה והרצה

### דרישות קדם
* מערכת הפעלה Windows.
* התקנה של [Node.js](https://nodejs.org/).
* Microsoft Outlook מוגדר במחשב.

### שלב 1: התקנת תלויות
פתח את הטרמינל בתיקיית הפרויקט והרצ:
```bash
npm install
שלב 2: הרצת השרת
כדי שהממשק יוכל לתקשר עם ה-Outlook שלך, השרת חייב לרוץ ברקע:

Bash
node local-agent/server.js
שלב 3: פתיחת הממשק
פתח את הדפדפן בכתובת:
http://localhost:3000

🔄 הרצה קבועה ברקע (מומלץ)
כדי שלא תצטרך להשאיר טרמינל פתוח, ניתן להשתמש ב-PM2:

התקנה: npm install pm2 -g

הרצה: pm2 start local-agent/server.js --name outlook-tool

השרת ירוץ ברקע גם לאחר סגירת הטרמינל.

📁 מבנה הפרויקט
web/ - קבצי הממשק הגרפי (HTML, CSS, JS).

local-agent/ - שרת ה-Node.js וסקריפטי האוטומציה.

temp/ - תיקייה זמנית לקבצים המועלים (מנוקה אוטומטית).


---

### שלב אחרון - העלאה ל-GitHub:
אחרי ששמרת את הקובץ, הרץ את הפקודות האלו כדי שכולם יראו את ה-README היפה שלך:

```bash
git add README.md
git commit -m "Add professional README"
git push
