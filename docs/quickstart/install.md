# התקנה

ההתקנה נעשית פעם אחת במחשב, ולוקחת כדקה.

## שלב 1 — התקנת pyRevit

pyArea הוא תוסף שרץ בתוך pyRevit, אז צריך קודם להתקין אותו.

הורדו והתקינו את [pyRevit גרסה 6.5.0](https://github.com/pyrevitlabs/pyRevit/releases/tag/v6.5.0.26173%2B1406).

## שלב 2 — פתיחת שורת הפקודה

לחצו ++win+r++, הקלידו `cmd` ולחצו ++enter++.

## שלב 3 — התקנת pyArea

העתיקו את השורה הבאה, הדביקו בחלון שנפתח ולחצו ++enter++:

```
pyrevit extend ui pyArea "https://github.com/blliav/pyArea.git"
```

## שלב 4 — הפעלה מחדש של רוויט

סגרו את רוויט ופתחו אותו מחדש. אמורה להופיע לשונית חדשה בשם **pyArea**.

---

## עדכון לגרסה האחרונה

כדאי להריץ מדי פעם — התוסף מתעדכן תכופות:

```
pyrevit extensions update pyArea
```

## הסרת התוסף

```
pyrevit extensions delete pyArea
```

---

!!! tip ""
    אחרי ההתקנה, המשיכו לעמוד [החישוב הראשון שלך](first-calculation.md).
