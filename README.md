# Duct Order Management System

A professional Windows desktop application built with **C# WinForms + .NET** and **ClosedXML** for Excel storage.

---

## 📁 Project Structure

```
DuctOrderApp/
├── DuctOrderApp.csproj
├── Program.cs
├── Forms/
│   └── MainForm.cs          ← All UI (built fully in code, no Designer file)
├── Models/
│   └── OrderModel.cs        ← Data model
├── Services/
│   └── ExcelService.cs      ← Excel read/write via ClosedXML
├── Helpers/
│   └── ThemeHelper.cs       ← Dark-mode colours, fonts, styling helpers
└── Orders.xlsx              ← Auto-created at runtime
```

---

## ⚙️ Prerequisites

| Tool | Version |
|------|---------|
| Visual Studio 2019 / 2022 (Community is fine) | Latest |
| .NET Framework 4.8 **or** .NET 6/8 (Windows) | See note below |
| NuGet package restore (automatic) | — |

> **Modern .NET note:** To target .NET 6 or .NET 8, change the `<TargetFramework>` line in  
> `DuctOrderApp.csproj` from `net48` to `net6.0-windows` or `net8.0-windows`.

---

## 🚀 Getting Started

### Option A – Visual Studio

1. Open **Visual Studio**.
2. **File → Open → Project/Solution** → select `DuctOrderApp.csproj`.
3. NuGet packages restore automatically (ClosedXML 0.102.2).
4. Press **F5** or click **Start** to build and run.

### Option B – .NET CLI

```bash
# From the DuctOrderApp/ folder:
dotnet restore
dotnet build
dotnet run
```

### Option C – NuGet manual install (if restore fails)

```powershell
# Package Manager Console inside Visual Studio:
Install-Package ClosedXML -Version 0.102.2
```

---

## ✨ Features

| Feature | Details |
|---------|---------|
| **Add orders** | Client · Order Name · Duct Type · Urgent · Done |
| **Auto date** | Current date stamped on save |
| **Excel storage** | `Orders.xlsx` auto-created; rows appended on save |
| **Load on startup** | All existing rows loaded into the grid automatically |
| **Edit rows** | Select a row → click ✏ Edit → modify → Save |
| **Delete rows** | Select a row → click ✖ Delete → confirm |
| **Search** | Real-time filter by client, order name, or duct type |
| **Filter urgent** | Toggle "⚡ Urgent only" checkbox |
| **CSV export** | Export filtered data to `.csv` |
| **Stats dashboard** | Total · Urgent · Done · Pending counters |
| **Dark mode** | Full dark theme across all controls |
| **Input validation** | Prevents empty submissions |
| **Toast notifications** | Non-blocking green success banner |

---

## 🗄️ Excel File

`Orders.xlsx` is created in the same directory as the executable  
(`bin/Debug/net48/` during development).

| Column | Contents |
|--------|----------|
| CLIENT | Client name |
| ORDER NAME | Order description |
| DUCT TYPE | Type of duct |
| URGENT | YES / NO |
| DATE CREATED | yyyy-MM-dd |
| DONE | YES / NO |

---

## 🎨 Dark Mode Palette

| Element | Colour |
|---------|--------|
| Form background | `#232323` |
| Control background | `#2D2D2D` |
| Input fields | `#373737` |
| Accent (buttons/title) | `#1E88E5` Blue |
| Urgent highlight | `#E53935` Red |
| Done highlight | `#43A047` Green |
| Text | White / `#B4B4B4` |

---

## 🔮 Extending the App

- **PDF export** – add `iTextSharp` or `QuestPDF` NuGet package and wire up `BtnExport_Click`.
- **Email notifications** – use `System.Net.Mail.SmtpClient` in `ExcelService`.
- **Dark title bar** (Win 11) – P/Invoke `DwmSetWindowAttribute` with `DWMWA_USE_IMMERSIVE_DARK_MODE`.
