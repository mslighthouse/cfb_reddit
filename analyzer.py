import sqlite3
from openpyxl import Workbook

conn = sqlite3.connect('firstquarter.db')   # CHANGE DEPENDING ON DATABASE
c = conn.cursor()

# Create Excel file
wb = Workbook()
ws = wb.active
ws.title = "First Quarter"

# create titles
ws['A1'].value = "Username"
ws['B1'].value = "Posts"
ws['C1'].value = "Flair 1"
ws['D1'].value = "Flair 2"


prev_user = ""
post_count = 0
user_total = 0
delete_num = 0

# unique users (EXCLUDING [deleted]) & post count (INCLUDING [deleted])
for row in c.execute('SELECT username FROM comments ORDER BY username'):
    if row != prev_user and "/u/None" not in row:
        prev_user = row
        user_total = user_total + 1
    if "/u/None" in row:
        delete_num = delete_num + 1
    post_count = post_count + 1
print("There were " + str(user_total) + " total users contributing " + str(post_count) + " total comments.")
print("Of the " + str(post_count) + " total posts, " + str(delete_num) + " have been deleted.")

rown = 2 # row number iterator
# Users groups in order of highest post amount
for row in c.execute('SELECT username, count(*) FROM comments GROUP BY username ORDER BY username'):
    row_fmt = str(row)[3:]
    row_fmt = row_fmt[:-1]
    row_fmt = row_fmt.replace("\',", "")
    row_fmt = row_fmt.split() #0 is username, #1 is post count
    
    ws.cell(row=rown, column=1).value = row_fmt[0]
    ws.cell(row=rown, column=2).value = int(float(row_fmt[1]))
    rown = rown + 1

# FLAIRS
bama_fan = 0
clem_fan = 0
bamaclem = 0
sec_fan  = 0
acc_fan  = 0

bama_flair = ['Alabama Band', 'Crimson Tide']
sec_flair  = ['Crimson Tide', 'Kentucky Wildcats', 'Carolina Gamecocks', 'Arkansas Razorbacks'
                'LSU', 'Volunteer', 'Ole Miss', 'Auburn', 'Mississippi State', 'Texas A&M'
                'Florida Gators', 'Missouri Tigers', 'Vanderbilt', 'Georgia Bulldog', 'SEC']
acc_flair = ['Boston College', 'Georgia Tech', 'Carolina State Wolf', 'Virginia Tech', 'Clemson',
                'Louisville Card', 'Pittsburg Panthers', 'Wake Forest Demon', 'Duke Blue Devils',
                'Miami Hurricanes', 'Syracuse', 'Florida State', 'North Carlonia Tar', 'Virginia Caveliers', 'ACC']

rown = 2
for row in c.execute ('SELECT flair1, flair2 FROM comments GROUP BY username ORDER BY username'):
    flair = str(row)[3:]
    flair = flair[:-2]
    flair = flair.replace("', u'", "|")
    flair = flair.replace("\", u'", "|")
    flair = flair.replace("\', u\"", "|")
    flair = flair.split("|")
    flair1 = flair[0]
    flair2 = flair[1]
    
    ws.cell(row=rown, column=3).value = flair1
    ws.cell(row=rown, column=4).value = flair2
    rown = rown + 1
    
    # Check Bama Fan
    if any(bama in flair1 for bama in bama_flair) or any(bama in flair2 for bama in bama_flair):
        bama_fan = bama_fan + 1
    # Check Clemson Fan
    if 'Clemson' in flair1 or 'Clemson' in flair2:
        clem_fan = clem_fan + 1
    # Check if Bama / Clemson or Clemson / Bama
    if ('Clemson' in flair1 and any(bama in flair2 for bama in bama_flair)) or any(bama in flair1 for bama in bama_flair) and 'Clemson' in flair2:
        bama_fan = bama_fan - 1
        clem_fan = clem_fan - 1
        bamaclem = bamaclem + 1

    # Check SEC flairs
    if any(sec in flair1 for sec in sec_flair) or any(sec in flair2 for sec in sec_flair):
        sec_fan = sec_fan + 1
    # Check ACC flairs
    if any(acc in flair1 for acc in acc_flair) or any(acc in flair2 for acc in acc_flair):
        acc_fan = acc_fan + 1

wb.save('first_results.xlsx')

print("Bama fans: " + str(bama_fan))
print("Clemson fans: " + str(clem_fan))
print("Bastards: " + str(bamaclem))
print("SEC Flairs: " + str(sec_fan))
print("ACC Flairs: " + str(acc_fan))