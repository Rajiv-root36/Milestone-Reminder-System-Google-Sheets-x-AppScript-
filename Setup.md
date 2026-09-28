# 🛠️ Setup Guide

This guide walks you through setting up the Milestone Reminder System from scratch.

---

## Step 1 — Copy the Google Sheet

1. Open the shared Google Sheet link provided by your team
2. Click **File → Make a copy**
3. Save it to your own Google Drive

---

## Step 2 — Add the Script

1. In your copied sheet, go to **Extensions → Apps Script**
2. Delete the default `myFunction` code
3. Paste the entire contents of `Code.gs` into the editor
4. Hit **Ctrl+S** to save
5. Close the Apps Script tab and **reload your Google Sheet (F5)**

You should now see a **"📋 Project Reminders"** menu in the top menu bar.

---

## Step 3 — Activate Your Reminders

1. Click **📋 Project Reminders → Activate daily email trigger**
2. Google will show a permissions popup — click through:
   - Click **"Advanced"**
   - Click **"Go to [project name] (unsafe)"**
   - Click **"Allow"**
3. ✅ A welcome confirmation email will arrive in your inbox

> This is a one-time step. You will never need to do this again.

---

## Step 4 — Add Your First Project

In the Dashboard sheet, fill in a new row:

| Column | What to enter |
|--------|--------------|
| **Project Name** | e.g. `SWISSPORT - Buyer` |
| **Start Date** | Click the cell — a date picker will appear |
| **End Date** | Click the cell — a date picker will appear |
| **Complexity** | Select from dropdown: `FT`, `C1`, `C2`, or `C3` |

All 9 milestone dates will **calculate automatically** as soon as you fill in all four columns.

---

## Step 5 — Sit Back

Every morning at **8 AM**, you will receive an email listing every task due that day across all your active projects. No further action needed.

---

## 📝 Notes

- **Status** is auto-managed — you don't need to change it manually
- **On Hold** is the only status you set manually — the system will never auto-change it
- **Manual date override** — you can click any milestone date cell and type a custom date. Your override will be preserved unless you edit the Start Date, End Date, or Complexity
- If no tasks are due on a given day, no email is sent

---

## ❓ Troubleshooting

| Problem | Solution |
|---------|----------|
| Menu not showing | Reload the sheet (F5) after saving the script |
| Not receiving emails | Check spam/junk folder; re-run Activate trigger |
| Dates not calculating | Make sure all 4 columns are filled (Name, Start, End, Complexity) |
| Wrong dates showing | Check that start date is a Monday for FT projects |
