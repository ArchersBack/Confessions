Named INFO.md because if it was README.md github gists has a bug that changes the title name of the gist to README.md

Recommended usage is as follows:

Place the Microsoft Office(`xlsx`) (not Microsoft Office Legacy(`xls`), those don't work as well or may not at all, don't know) and this script in the same folder. If you have an OAuth2Util oauth.ini, put that in here to. The script will attempt to use OAuth if it can, else it will default to the username ad password arguments, if possible, else raise an error.

To run, either:

Open python3 in a terminal, depending on the OS it's either `python3` or `py -3`.
Type

    >>> from modapplicationposter import run
    >>> run('nameofworkbookwithoutxlsx', 'columnletterforusernamesifapplicable',
    ... subreddit='subreddityouarepostingtowithout/r/' **kw)
    
You can also do this in a terminal/cmd:

    python3 -c "from modapplicationposter import run; run('nameofworkbookwithoutxlsx', 'columnletterforusernamesifapplicable', subreddit='subreddityouarepostingtowithout/r/' **kw)"
    
Note: Like above, you may need to replace `python3` with `py -3`

`**kw` are any of the other params as mentioned by the docstring in modapplicationposter.py.
