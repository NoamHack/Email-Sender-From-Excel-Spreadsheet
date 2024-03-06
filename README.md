# תיעוד לסקריפט שליחת דוא"ל באמצעות גיליון אלקטרוני

סקריפט Python זה מיועד לשליחת דוא"ל אישיים לרשימת נמענים המוגדרת בגיליון אלקטרוני, תוך שימוש בספריית `pandas` לניהול הגיליון, `smtplib` לשליחת הדוא"ל דרך פרוטוקול SMTP, ו-`email.mime.text` MIMEText לעיצוב תוכן הדוא"ל. הסעיפים הבאים מפרטים את השלבים הנדרשים להכנה, הבנה והרצה של הסקריפט על מכונה מקומית.

## מדריך הכנה והרצה:

**התקנת Python**: וודאו ש-Python מותקן במערכת שלכם (מומלץ גרסה 3.x). Python מהווה את הבסיס להרצת הסקריפט.

**התקנת תלותים**: הסקריפט דורש את ספריית ה-`pandas` לניהול נתוני גיליון אלקטרוני. התקינו אותה באמצעות pip אם טרם עשיתם זאת:
    ```bash
    pip install pandas
    ```

   **הכנת גיליון אלקטרוני**: הכינו גיליון אלקטרוני של Excel המכיל לפחות שתי עמודות בשמות `Name` ו-`Email` לשם ולכתובת הדוא"ל של הנמען, בהתאמה. שמרו את הגיליון בתיקייה בשם `input` באותו המדריך עם הסקריפט, והחליפו את `YOUR_FILE.xlsx` בקוד לשם הקובץ שלכם.

   **הכנת חשבון הדוא"ל**: שנו את המשתנים `your_email` ו-`your_password` בסקריפט עם כתובת הדוא"ל והסיסמה של השולח. לביטחון מוגבר, במיוחד עם Gmail, מומלץ ליצור ולהשתמש בסיסמת אפליקציה.

   **הרצה**: לאחר שהכנה הושלמה, הריצו את הסקריפט מהטרמינל או משורת הפקודה שלכם:
    ```bash
    python <script_name>.py
    ```
    
    החליפו את `<script_name>` בשם הקובץ המדויק של הסקריפט שלכם.

## סקירה כללית של התפקוד:

הסקריפט פועל דרך מספר שלבים עיקריים:

   **חיבור לשרת SMTP**: יוזם חיבור מאובטח לשרת ה-SMTP של Gmail באמצעות SSL בפורט 465 ומנסה לאמת עם האישורים המסופקים. אימות תקין הוא קריטי להמשך פעולות שליחת הדוא"ל.

   **עיבוד גיליון**: משתמש ב-`pandas` לקריאת קובץ Excel המצוין, ומוציא שמות וכתובות דוא"ל של הנמענים כדי ליצור את רשימת הנמענים. נתיב הקובץ הנכון והפורמט הם קריטיים.

   **הרכבה ושליחת הדוא"ל**: מעבד באופן איטרטיבי כל רשומה ברשימת הנמענים, מרכיב דוא"ל אישי לכל נמען באמצעות שמו. מגדיר נושא לדוא"ל, מציין את השולח והנמען ומנסה לשלוח את הדוא"ל דרך החיבור ל-SMTP שנוצר.

   **טיפול בשגיאות**: כולל טיפול בשגיאות של כשלונות באימות SMTP, חריגות של קובץ לא נמצא, וחריגות כלליות בקריאת קובץ ה-Excel. מבטיח שחיבור השרת SMTP נסגר בחן לפני סיום ההרצה של הסקריפט במקרים של שגיאות. 
