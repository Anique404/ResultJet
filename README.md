# 🚀 RESULTJET

**School Result Card Automation System | KG to 10th | Excel to PDF | 33% Pass Rule 

---

## 📌 WHAT IS RESULTJET?

ResultJet is a desktop app that **automatically generates result cards** from Excel files. Built for Pakistani schools (KG to 10th).

**One click = 100+ result cards ready in 2 minutes**

---

## ⚡ KEY FEATURES

| Feature | Description |
|--------|-------------|
| 🎯 **Classes** | KG, 1st to 10th (11 classes) |
| 📊 **Excel Import** | Auto-detect subjects & max marks |
| 🧮 **33% Rule** | FAIL if ANY subject <33% |
| 🎨 **PDF Cards** | Professional single template |
| 🗄️ **Database** | SQLite - no setup needed |
| 🖥️ **GUI** | Simple dropdown & click |
| ⚡ **Speed** | 100 PDFs in <90 seconds |

---

## 📁 FILE STRUCTURE

```
ResultJet/
├── ResultJet.exe        # Double click to run
├── CLASS_DATA/          # Place your Excel files here
├── OUTPUT/             # PDFs generated here
└── DATABASE/           # Student data stored here
```

---

## 📊 EXCEL FORMAT

| RollNo | StudentName | FatherName | English | English_Max | Math | Math_Max |
|--------|------------|-----------|---------|-------------|------|----------|
| 1 | Ali Raza | Raza Ahmed | 45 | 50 | 20 | 50 |
| 2 | Sara Khan | Khan Sahab | 48 | 50 | 45 | 50 |

**Rules:**
- First 3 columns FIXED: `RollNo`, `StudentName`, `FatherName`
- Each subject needs TWO columns: `Subject` + `Subject_Max`
- Max marks can be different (50, 75, 100...)

---

## 🎯 PASS/FAIL RULE

```
✅ PASS = ALL subjects ≥ 33%
❌ FAIL = ANY subject < 33%
```

**Grade System:**
- 90%+ = A+
- 80-89% = A
- 70-79% = B
- 60-69% = C
- 50-59% = D
- 33-49% = E
- <33% = F (FAIL)

---

## 🖥️ HOW TO USE 

**STEP 1:** Put Excel files in `CLASS_DATA` folder
**STEP 2:** Double click `ResultJet.exe`
**STEP 3:** Select Class → Click "Generate PDFs"
**STEP 4:** Get PDFs from `OUTPUT` folder

**Done!** ✅

---

## 📥 INSTALLATION

```
1. Download ResultJet.zip
2. Extract to Desktop
3. Double click ResultSystem.py
4. Start using!
```

**No Python. No installation. Just double click.**

---

## 📸 SAMPLE FILES INCLUDED

```
SAMPLE_FILES/
├── KG_Sample.xlsx      (Test with KG class)
├── 5th_Sample.xlsx     (Test with 5th class)
└── 10th_Sample.xlsx    (Test with 10th class)
```

---

## ❗ QUICK TROUBLESHOOTING

| Problem | Solution |
|--------|----------|
| "No Excel files" | Place files in CLASS_DATA folder |
| "No _Max column" | Add Subject_Max columns |
| "Marks exceed max" | Check marks ≤ max marks |
| "PDF not generating" | Check OUTPUT folder |

---

## ✅ PERFECT FOR:

- 🏫 **School Principals**
- 👨‍🏫 **Class Teachers**
- 📚 **Exam Controllers**
- 🎓 **Small/Medium Schools**

---

## 📞 SUPPORT

**Email:** support@resultjet.com  
**WhatsApp:** +92 300 1234567  

---

## 📜 LICENSE

MIT License - Free for all schools

---

## ⭐ STAR THIS REPO

If ResultJet helps your school, give it a star! ⭐

---

**🇵🇰 Made in Pakistan for Pakistani Schools**

**[DOWNLOAD NOW](https://github.com/yourusername/ResultJet/releases)**

---

🚀 **ResultJet - Fast. Accurate. Professional.**
