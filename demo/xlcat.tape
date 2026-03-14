# VHS tape for recording xlcat demo GIF.
# Usage: vhs demo/xlcat.tape
#
# Prerequisites:
#   - xlcat binary built and in PATH
#   - Demo xlsx files created: cargo run --example create_demo

Output demo/xlcat.gif

Set FontSize 18
Set Width 1400
Set Height 700
Set Padding 20
Set Theme "GruvboxDarkHard"

Set TypingSpeed 80ms

Set Shell "bash"

Hide
    Type 'export PS1="> "'
    Enter
    Type "clear"
    Enter
Show

# ─── 1. VIEW A SPREADSHEET ───
Type "xlcat demo/sales.xlsx"
Enter
Sleep 3s

# ─── 2. SCHEMA ONLY ───
Type "xlcat demo/sales.xlsx --schema"
Enter
Sleep 2s

# ─── 3. SUMMARY STATISTICS ───
Type "xlcat demo/sales.xlsx --describe"
Enter
Sleep 3s

# ─── 4. MULTI-SHEET WORKBOOK ───
Type "xlcat demo/budget.xlsx"
Enter
Sleep 3s

# ─── 5. SELECT A SHEET ───
Type "xlcat demo/budget.xlsx --sheet Revenue"
Enter
Sleep 2s

# ─── 6. HEAD AND TAIL ───
Type "xlcat demo/sales.xlsx --head 3 --tail 2"
Enter
Sleep 2s

# ─── 7. CSV OUTPUT ───
Type "xlcat demo/sales.xlsx --csv --head 5"
Enter
Sleep 3s
