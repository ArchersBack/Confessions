BAD_WORDS = ['shit', 'piss', 'fuck']

def run(workbookname, usernamecolletter=None, subreddit=None, notifyme=False, useragent=None, username=None, password=None):
    """
    Main instance

    The user that is OAuthed or LoginAuthed must be an approved contributor to the sub to post apps
    and must be an approved submitter to the sub. Just making it a mod won't work.

    :param usernamecolletter: the letter of the column that denotes the applicant's username.
    If not specified, it will make the title have the timestamp instead of the user's name. This
    script assumes that the timestamp row is the first row, so it uses appropiate grammar to comply.
    :param notifyme: whether or not to send inbox replies to the account that is making the posts. default is False.
    :param subreddit: the sub to post to (string or praw.objects.Subreddit object). While it is a
    keyword argument, that's only cause technicalities. IT IS REQUIRED.
    :param useragent: optional useragent to specify. The default is
    'Mod application to subreddit poster posting to /r/{subreddit} by /u/13steinj'
    :params username, password: A username and password to specify if you wish to use LoginAuth.
    If OAuth is available, OAuth will have priority. If it fails, login auth will be used.
    IN PRAW4 LOGIN AUTH WILL BE DEPRECATED, JUST AS A REMINDER IN CASE YOU SAY THIS SCRIPT DOESN'T WORK WHEN IT'S OUT
    
    :returns: The amount of posts made and the sub they were made to.
    """
    if subreddit == None:
        raise TypeError("run() missing required keyword argument subreddit, dipshit")
    import praw
    import openpyxl
    useragent = (useragent or 
                 ("Mod application to subreddit "
                  "poster posting to /r/{0} by /u/13steinj".format(subreddit)))
    r = praw.Reddit(useragent)
    try:
        import OAuth2Util
        o = OAuth2Util.OAuth2Util(r)
        o.refresh(force=True)
    except Exception as e:
        if e.__class__ == ImportError and username and password:
            print("OAuth2Util not installed, using login")
            r.login(username, password)
        elif username and password:
            print("OAuth2Util failure, attempting to use login")
            r.login(username, password)
            print("Success")
        else:
            raise
    workbook = openpyxl.load_workbook('{0}.xlsx'.format(workbookname))
    sheets = workbook.get_sheet_names()
    if len(sheets) > 1:
        sheetlist = '\n'.join(['{0}: {1}'.format((sheets.index(i)+1), i) for i in sheets])
        num = input("Which sheet would you like?\n{0}\n".format(sheetlist))
        try:
            num = int(num)
            if num <= 0:
                raise IndexError("What a dipshit this user is. If only we could replace them with more code...")
            sheet = sheets[num-1]
        except ValueError:
            raise ValueError("You had to choose a number, dipshit")
        except IndexError:
            raise IndexError("That sheet doesn't exist, dipshit.")
    else:
        sheet = sheets[0]
    spreadsheet = workbook[sheet]
    rows = list(spreadsheet.rows)
    questionrow = rows.pop(0)
    questionvals = {cell.column: cell.value for cell in questionrow if cell.value is not None}
    if usernamecolletter and usernamecolletter not in questionvals:
        raise KeyError("That usernamecolumnletter doesn't exist, dipshit")
    rownum = 0
    for row in rows:
        # clear out username, title, and body for next run
        rownum += 1
        username = None
        body = ''
        # make body
        for cell in row:
            if cell.column not in questionvals:
                continue
            body += "#{0}\n\n".format(questionvals[cell.column])
            body += "{0}\n\n---\n\n".format(cell.value if cell.value else "*No Answer*")
            if username:
                continue
            else:
                username = cell.value if cell.column == usernamecolletter else None
        titleinfo = "at {0}".format(row[0].value) if not username else "by {0}".format(username)
        title = "Moderator Application #{0} {1}".format(rownum, titleinfo)
        post = r.submit(subreddit, title, text=body, send_replies=notifyme)
        print("Submitted {0}".format(post.permalink))
    print("I've just posted {0} applications, dipshit. I'm done now. Call me when you want me again. But you could at least ask me out to dinner first, you rascal!".format(rownum))
    return rownum, subreddit
    
def background_check(reddit_session, username, post_sub=None, badwords=[]):
    from praw.errors import NotFound, Forbidden
    import datetime
    cannot_submit = reddit_session.user == None and not reddit_session.is_oauth_session()
    if cannot_submit and post_sub != None:
        raise ValueError("Can't post without being logged in, dipshit")
    title = "Background Check: {0}".format(username) if post_sub else None
    body = "Background Check:\n\n" if not title else ""
    user = reddit_session.get_redditor(username)
    try:
        dumblist = list(user.get_overview())
    except NotFound:
        body += "Account doesn't exist"
        return body
    except Forbidden:
        body += "Account deleted or shadowbanned or permanently suspended"
        return body
    ucreated = datetime.datetime.fromtimestamp(user.created_utc)
    utimeago = "Account made {0} ago".format(str(datetime.datetime.now() - ucreated))
    utimeon = "on " + str(ucreated)
    utimetext = " ".join([utimeago, utimeon])
    body += utimetext + "\n\n"
    posts = list(user.get_submitted(limit=None))
    try:
        timepost = posts[100]
        timepostnum = 100
        timepostord = ordinal(timepostnum)
        timepostpre = timepostord + " post"
        timepostdate = datetime.datetime.fromtimestamp(timepost.created_utc)
        timeago = "made {0} ago".format(str(datetime.datetime.now() - timepostdate))
        timeon = "on " + str(timepostdate)
        timeposttext = " ".join([timepostpre, timeago, timeon])
    except IndexError:
        try:
            timepost = posts[-1]
            timepostnum = (posts.index(timepost) + 1)
            timepostord = ordinal(timepostnum)
            timepostpre = "last({0}) post".format(timepostord)
            timepostdate = datetime.datetime.fromtimestamp(timepost.created_utc)
            timeago = "made {0} ago".format(str(datetime.datetime.now() - timepostdate))
            timeon = "on " + str(timepostdate)
            timeposttext = " ".join([timepostpre, timeago, timeon])
        except IndexError:
            timeposttext = "User has not made a single post, or all are deleted"
    body += timeposttext + "\n\n"
    comments = list(user.get_comments(limit=None))
    totalprofanities, specificprofanities = profanitycheck(badwords, posts, comments)
    table = "Profanity | Times Used\n---|---"
    for word in specificprofanities:
        table += "\n"
        table += "{0} | {1}".format(word, specificprofanities[word])
    if specificprofanities:
        body += "Profanities used in the last 1000 posts and comments"
        body += table

def ordinal(value):
    """
    Converts zero or a *postive* integer (or their string 
    representations) to an ordinal value.

    http://code.activestate.com/recipes/576888-format-a-number-as-an-ordinal/

    """
    try:
        value = int(value)
    except ValueError:
        return value

    if value % 100//10 != 1:
        if value % 10 == 1:
            ordval = u"%d%s" % (value, "st")
        elif value % 10 == 2:
            ordval = u"%d%s" % (value, "nd")
        elif value % 10 == 3:
            ordval = u"%d%s" % (value, "rd")
        else:
            ordval = u"%d%s" % (value, "th")
    else:
        ordval = u"%d%s" % (value, "th")

    return ordval

def profanitycheck(badwords, *args):
    """
    Profanity Checker
    :param badwords: list of badwords
    :param args: comma delimited list of lists of posts or comments
    :returns: dict of word: usagecount for each word in badwords
    """
    import re
    from bs4 import BeautifulSoup
    obj_list = []
    html_list = []
    check_list = []
    profanity_usage = {}
    for l in args:
        obj_list += l
    html_list += [obj.selftext_html for obj in obj_list if hasattr(obj, 'selftext_html')]
    html_list += [obj.body_html for obj in obj_list if hasattr(obj, 'body_html')]
    for html in html_list:
        if html == None:
            continue
        check_list += BeautifulSoup(html).get_text().split()
    badwordtotalusage = 0
    for word in badwords:
        badwordusage = 0
        for check in check_list:
            if re.search(word, check):
                badwordusage +=1
        profanity_usage[word] = badwordusage
        badwordtotalusage += badwordusage
    return badwordtotalusage, profanity_usage