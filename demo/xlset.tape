# VHS tape for recording xlset demo GIF.
# Usage: vhs demo/xlset.tape
#
# Prerequisites:
#   - xlcat and xlset binaries built and in PATH
#   - Demo xlsx files created: cargo run --example create_demo

Output demo/xlset.gif

Set FontSize 14
Set Width 1100
Set Height 550
Set Padding 15
Set Theme "GruvboxDarkHard"

Set TypingSpeed 80ms

Set Shell "bash"

Hide
    Type 'export PS1="> "'
    Enter
    Type "cp demo/sales.xlsx /tmp/demo_edit.xlsx"
    Enter
    Type "clear"
    Enter
Show

# ─── 1. VIEW BEFORE EDITING ───
Type "xlcat /tmp/demo_edit.xlsx --head 3"
Enter
Sleep 2s

# ─── 2. SET A SINGLE CELL ───
Type "xlset /tmp/demo_edit.xlsx B2=South"
Enter
Sleep 2s

# ─── 3. SET MULTIPLE CELLS ───
Type "xlset /tmp/demo_edit.xlsx D2=99999 E2=500"
Enter
Sleep 2s

# ─── 4. VERIFY CHANGES ───
Type "xlcat /tmp/demo_edit.xlsx --head 3"
Enter
Sleep 3s

# ─── 5. TYPE TAG: PRESERVE LEADING ZERO ───
Type "xlset /tmp/demo_edit.xlsx C2:str=07401"
Enter
Sleep 2s

Type "xlcat /tmp/demo_edit.xlsx --head 2"
Enter
Sleep 2s

# ─── 6. BULK UPDATE FROM CSV ───
Hide
    Type "printf 'cell,value\nA5,2024-06-15\nB5,HQ\nD5,50000\n' > /tmp/updates.csv"
    Enter
    Sleep 300ms
Show

Type "cat /tmp/updates.csv"
Enter
Sleep 2s

Type "xlset /tmp/demo_edit.xlsx --from /tmp/updates.csv"
Enter
Sleep 2s

# ─── 7. WRITE TO NEW FILE ───
Type "xlset /tmp/demo_edit.xlsx --output /tmp/final.xlsx A1=ID"
Enter
Sleep 2s

Type "xlcat /tmp/final.xlsx --head 3"
Enter
Sleep 3s
