import praw
import time
import sqlite3

start = time.time()
print("Program Starting") # for debugging / timekeeping

# connect to sql database
conn = sqlite3.connect('cfb.db')
cursor = conn.cursor()

# reddit authorization / identification
reddit = praw.Reddit("CFB Natty Comment Scrapper 1.0 by /u/mslighthouse")

# grab submission (thread)
submission = reddit.get_submission(submission_id='40jtc4')      # CHANGE DEPENDING ON THREAD
submission.replace_more_comments(limit=None, threshold=0)

# flatten tree of all comments
commentlist = praw.helpers.flatten_tree(submission.comments)

for comment in commentlist:
    
    #author
    auth = comment.author
    if auth is None:    # tests for [deleted] author names
        auth = "None"
    author = '/u/' + str(auth)
    
    # flair(s) if applicable
    flair_text = comment.author_flair_text
    if flair_text is None:  # tests for no flairs (seriously? REPRESENT, C'MON MAN)
        flair1 = "None"
        flair2 = "-"
    else:
        if '/' in flair_text:   # if two flairs
            flair1 = flair_text.split('/')[0].strip()
            flair2 = flair_text.split('/')[1].strip()
        else:
            flair1 = flair_text
            flair2 = "-"

# comment bodies
if comment.body is None:
    body = "None"
    else:
        body = comment.body.encode('ascii', 'ignore')

# add to SQL Database CHANGE NAME ACCORDING TO DATABASE
cursor.execute("INSERT INTO first_quarter VALUES (?,?,?,?);", (author, flair1, flair2, body))

# commit changes
conn.commit()

# timekeeping
print("Program ended. Process time:" + str((time.time() - start)) )