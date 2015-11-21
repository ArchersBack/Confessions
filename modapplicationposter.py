def run(workbookname, subreddit, useragent=None, usernamecol=None, notifyme=True):
    """
    Main instance

    :param usernamecol: the text in the column that denotes the applicant's username
    :param notifyme: whether or not to send inbox replies to the account that is making the posts
    """
    import praw
    import OAuth2Util
    import openpyxl
    useragent = (useragent or 
                 ("Mod application to subreddit "
                  "poster posting to /r/{0}".format(subreddit)))
    r = praw.Reddit(useragent)
    o = OAuth2Util.OAuth2Util(r)
    o.refresh(force=True)
    spreadsheet = openpyxl.load_workbook('{0}.xlsx'.format(workbookname))
    sheets = spreadsheet.get_sheet_names()
    if len(sheets) > 1:
        sheetlist = '\n'.join(['{0}: {1}'.format((sheets.index(i)+1), i) for i in sheets]))
        num = input("Which sheet would you like?\n{0}\n".format(sheetlist)
        try:
            num = int(num)
            if num <= 0:
                raise IndexError('What a dipshit this user is. If only we could replace them with more code...')
            sheet = sheets[num-1]
        except ValueError:
            raise ValueError('You had to choose a number, dipshit')
        except IndexError:
            raise IndexError('That sheet doesn't exist, dipshit.')
    else:
        sheet = sheets[0]
    spreadsheet = spreadsheet[sheet]
    rows = spreadsheet.rows
    questionrow = rows.pop(0)
    questionvals = {cell.column: cell.value for cell in questionrow}
    try:
        usernamecolnum = {v: k for k, v in questionvals.iteritems()}[usernamecol] if usernamecol else None
    except KeyError:
        raise KeyError('That usernamecolumn doesn't exist, dipshit')
    rownum = 0
    for row in rows:
        rownum += 1
        body = ''
        for cell in row:
            body += "#{0}\n\n".format(questionvals[cell.column])
            body += "{0}\n\n---\n\n".format(cell.value)
        titleinfo = "at {0}".format(row[1]) if not usernamecolnum else "by {0}".format(row[usernamecolnum])
        title = "Moderator Application #{0} {1}".format(rownum, titleinfo)
        r.submit(subreddit, title, text=body, send_replies=notifyme)
    print("I've just posted {0} applications, dipshit. I'm done now. Call me when you want me again. But you could at least ask me out to dinner first, you rascal!".format(rownum))
    return rownum, subreddit