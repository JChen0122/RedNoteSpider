import xlsxwriter, xlrd
import os, sys
import shutil
from datetime import datetime, timedelta
import time, random
# pipinstall DrissionPage, instructions https://drissionpage.cn/
from DrissionPage import ChromiumOptions, WebPage
import requests
import webbrowser


# recreate a file
def folder_to_save (p):
    if os.path.exists(p):
        print('检测到"{}"文件夹，为避免重复写入数据将自动删除'.format(p))
        try:
            shutil.rmtree(p)
        except:
            input('{}文件夹下可能有打开的文件，请关闭后再运行程序，任意键退出'.format(p))
            exit()
    if not os.path.exists(p):
        os.makedirs(p)

# save data to excel
def excel_write(ph, data):
    from math import ceil
    # create a table and write in data
    table = xlsxwriter.Workbook(ph)
    # avoid the rows limit for xlsx:1048576\xls:65536
    table_type = ph.split('.')[-1]
    if table_type == 'xlsx':
        # ceil: 向上取整
        sheet_num = ceil((len(data)-1)/1048000)
    else:
        sheet_num = int(len(data)/6500)
    print(f'* Total Data:{len(data)} rows; Target Sheet:{sheet_num} sheets\n* writing......')
    for s in range(sheet_num):
        sheet = table.add_worksheet(f'sheet{s}')
        if table_type == 'xlsx':
            id_begin = s * 1048000
            id_end = 1048000 + id_begin
        else:
            id_begin = s * 6500
            id_end = 6500 + id_begin
        if id_end > len(data):
            id_end = len(data)-1
        # write head
        sheet.write(0, 0, 'ID')
        for k in range(len(data[0])):
            sheet.write(0, k + 1, data[0][k])
        # write data
        for i in range(id_begin, id_end):
            # write ID column
            sheet.write(i-id_begin+1, 0, i + 1)
            for j in range(len(data[i + 1])):
                # 这是最宽泛的将数字、字符串分开的方法，不过有可能出错，可以考虑结合“判断是否为数字”的代码使用
                try:
                    sheet.write(i-id_begin+1, j + 1, data[i + 1][j])
                except:
                    sheet.write(i-id_begin+1, j + 1, str(data[i + 1][j]))
        print(f'* Have finished:sheet{s}')
    table.close()
    print('###save as ' + ph)

# file must existence
def exist(p):
    if not os.path.exists(p):
        input('"{}"文件不存在，请确认，任意键退出'.format(p))
        exit()

def input_check(must_in, l, f, count):
    # three chances
    for i in range(4):
        if i == 3:
            input('请重新运行脚本，任意键退出')
            exit()
        r = input()
        # numbers limit
        if len(r) == 0:
            print('输入值为空，剩余输入次数为:', (2 - i))
            continue
        elif len(r) > l:
            print('输入值超出限制，剩余输入次数为:', (2 - i))
            continue
        else:
            rr = r.split(',')
            # avoid same number
            g = 'out'
            rrr = []
            for j in rr:
                # must fit the set mode
                for k in must_in:
                    if int(j) == k:
                        g = 'in'
                        print(j)
                        rrr.append(j)
                        break
                    else:
                        g = 'out'
            if len(rrr) != count:
                g = 'out'
            if g == 'out':
                print('输入错误，剩余输入次数为:', (2 - i))
                continue
            else:
                for m in rrr:
                    print('input:', m)
                    f.append(int(m))
                break

def human_choose(choices, times):
	# choices: 供选择的list
	# times： 需要选择的个数
    id = []
    for i, j in enumerate(choices):
        print(i, ':', j)
        id.append(i)
    print(f'本次需选择{times}个数据，请输入对应编号：\n（严格按照目标字段顺序先后输入，英文逗号分隔，Enter确定）')
    file = []
    input_check(id, len(id), file, times)
    return file

# open a xlsx file and read data
def excel_read(p, select, f):
    # pip install xlrd==1.2.0
    exist(p)
    print(f'打开：{p}')
    data = xlrd.open_workbook(p)
    table = data.sheets()[0]
    lines = table.nrows
    # select specific columns
    if len(select) != 0:
        head = table.row_values(0)
        select_c = []
        for i, j in enumerate(head):
            if j in select:
                select_c.append(i)
                if len(select_c) == len(select):
                    break
        if len(select_c) != len(select):
            print(f'文件{p}中未找到目标字段{select}，请确认！\n')
            select_c = human_choose(head, len(select))
            print('选择列：', select_c)
        f.append(select)
        for i in range(lines-1):
            row_s = []
            for j in select_c:
                row_s.append(table.cell_value(i+1, j))
            f.append(row_s)
    else:
        for i in range(lines):
            f.append(table.row_values(i + 1))
    print(f'共读取到：{len(f)} 行数据（含首行）')
    return

def txt_read(path):
    """read txt file and return a list"""
    exist(path)
    f = open(path, mode='r', encoding='utf8')
    l = []
    wrong = 0
    for i in f.readlines():
        j = i.strip('\n')
        if len(j) != 0:
            l.append(j)
        else:
            wrong += 1
    print('* Have read the file: {}; empty lines : {}'.format(path, wrong))
    return l

# write data into txt
def txt_write(path, data):
    f = open(path, mode='a', encoding='utf8')
    data_line = []

    def list_split(d):
        if type(d) in (list, tuple):
            for m in d:
                list_split(m)
        else:
            data_line.append(d)
        return

    if type(data) in (list, tuple):
        for i in data:
            list_split(i)
            j = ''
            for k in data_line:
                j += ','+str(k)
            f.write(j.lstrip('"",')+'\n')
            data_line.clear()
    else:
        f.write(data)
    print('保存至：'+path)
    f.close()

def is_number(s):
        try:
            float(s)
            return True
        except ValueError:
            pass
        return False

# pretend to be human not machine
def human():
    import time
    from datetime import datetime, timedelta
    import random

    def trans(t1):
        # from str to datetime
        t2 = datetime.strptime(t1, '%H:%M')
        return t2

    def trans2(t3):
        # select %H:%M part from datetime
        t3 = t3.strftime('%H:%M')
        t4 = datetime.strptime(t3, '%H:%M')
        return t4

    tn = datetime.now()
    # beyond 5 minutes take a nap
    pause1 = nap['nap'] + timedelta(minutes=random.randint(5, 10))
    # beyond 0.5 hour have a sleep
    pause2 = nap['sleep'] + timedelta(minutes=random.randint(30, 60))
    # during lunch or dinner stop here
    pause3 = [('12:00', '12:30'), ('17:30', '18:00')]
    # sleep time
    ts = 0
    mode = ''
    # whether it's time to eat something
    for i in pause3:
        if trans(i[0]) < trans2(tn) < trans(i[1]):
            ts = random.uniform(3, 5)
            mode = 'dining'
            break
        else:
            mode = 'normal'
    if mode == 'normal':
        if tn < pause1:
            ts = random.uniform(3, 20)
        elif pause1 < tn < pause2:
            ts = random.uniform(30, 120)
            nap['nap'] = tn
        elif tn > pause2:
            ts = random.uniform(60, 180)
            nap['sleep'] = tn
    # whether it's late at night
    if trans('0:00') < trans2(tn) < trans('7:00'):
        mode = 'Late at night'
        ts = ts * tn.hour / 2
    # pause
    print('……{}:pause（{}s）……{}'.format(mode, round(ts, 2), tn))
    time.sleep(ts)

def log_in_check():
    def log_in_tbc():
        print('* checking log in status...')
        # main page + note page
        j = [('@class=reds-modal reds-modal-open login-modal', '登录'), ('@class=comments-login', '登录')]
        for i in j:
            try:
                m = page.ele(i[0]).text.find(i[1])
                break
            except:                
                m = -1
        return m

    i = log_in_tbc()
    while i != -1:
        print('* please log in manually..')
        check = ''
        while check != 'done':
            check = input('* 如问题解决，请输入“done”继续数据抓取\n输入“exit”退出程序')
            print('* input:',check)
            if check == 'exit'        :
                exit()
        i = log_in_tbc()
    print('* the user has logged in. Ready to continue.')

# end remark
def end():
    print('\n'
          '\n****************************'
          '\n*数据已全部获取，Enter键退出*'
          '\n****************************\n\n')
    input()
    exit()

def keywords_load():
    print('\n* Load keywords for searching...')
    print('* You can try searching on the website yourself and come back later')
    keywords_choice = ['load from text file', 'type in']
    keywords_choice = human_choose(keywords_choice, 1)[0]
    keywords = []
    keywords_path = 'keywords.txt'
    if keywords_choice == 0:
        # load from text file
        while os.path.exists(keywords_path) == False:        
            f = open(keywords_path, 'a')
            f.write('')
            f.close
            print('* please type in all the keywords (one each line) \n and save them in {} (just created in the save folder as this script)'.format(keywords_path))
            input('* Please type Enter when {} is ready...'.format(keywords_path))    
        keywords = txt_read(keywords_path)
        while len(keywords) == 0:
            print('* {} is empty'.format(keywords_path))
            input('* Please type Enter when {} is ready...'.format(keywords_path))
            keywords = txt_read(keywords_path)
    else:
        # load from typing in
        print('* How many keywords do you want to type in now?\
            \n** You can type in one keyword first and type in another after checking the search result (not recommended)')
        keywords_num = input()
        if is_number(keywords_num):
            for i in range(int(keywords_num)):
                j = input('keyword {} of {}:'.format(i+1, int(keywords_num)))
                print('* input: ', j)
                keywords.append(j)
            txt_write(keywords_path, keywords)
    print('* load keywords:', keywords)
    return keywords

def search_keywords(keywords):
    # mode = together, afterwards
    # extract targeted note items for each scrolldown 
    # (due to limitation, only a certain number of notes will be loaded and kept each time)
    def save_notes(keyword):
        web_prefix = 'https://www.xiaohongshu.com'
        notes_temp = page.eles('@class=note-item')
        notes_new = 0    
        for i in notes_temp:
            try: 
                '''<ChromiumElement a data-v-30d73e1a='' href='/user/profile/5c66eb8600000000110130be?channel_type=web_search_result_notes&parent_page_channel_type=web_profile_board' class='author' target='_blank'> 
                浙江农村美丽小院出租 
                <ChromiumElement a data-v-30d73e1a='' href='/explore/66a10155000000000d00ff52' style='display: none;'>'''
                # note_url
                j = str(i.ele('@href^/explore/'))
                m = j.find('href=')
                n = j.find('\' style=')
                note_url = web_prefix+j[m+6:n]
                #  check duplicate and save data
                m = len(notes_url)
                notes_url.add(note_url)
                m2 = len(notes_url) 
                # if add a new note
                if  m2 > m: 
                    # new notes           
                    note_author = str(i.ele('@class=name').text)
                    note_title = i.ele('@class=title').text
                    # note_author_profile
                    j = str(i.ele('@class=author'))
                    m = j.find('href=')
                    n = j.find('\' class=')
                    note_author_profile = web_prefix+j[m+6:n]
                    search_time = str(datetime.today())
                    note = [note_author_profile, note_title, note_url, keyword, search_time]                    
                    if note_author not in notes.keys():
                        notes[note_author] = []
                    notes[note_author].append(note)
                    notes_new += 1
                    print('*** Find note:', note[1])
                    # one saved
                    search_result_all[keyword][0] += 1
                else:
                    # one duplicate
                    search_result_all[keyword][2] += 1
            except:
                pass
        return notes_new
    
    def extract_notes():
        # save preliminary data to excel
        print('** save all search results...')
        # save all data in one list
        notes_save =[]
        notes_save.append(notes['ID'][0])
        # {'author': [notes]} to [author, notes]
        for i in notes.keys():
            if i != 'ID':
                for j in notes[i]:
                    # author
                    m = [i]
                    # note
                    for n in j:
                        m.append(n)             
                    notes_save.append(m)
        excel_write(search_result, notes_save)
    
    # use notes_url to check duplicates
    notes_url = set()
    search_result_all = {'keyword': ['the amount saved', 'the amount loaded', 'the amount duplicated']}
    # dic to save notes data. one author can have different notes, ID = user_name (independant)
    notes = {'ID':[('author', 'author_profile', 'title', 'note_url - unaccessible', 'search keyword', 'search time')]}
    folder_to_save('data')
    for keyword_num in range(0, len(keywords)):
        keyword = keywords[keyword_num]
        search_result_all[keyword] = [0, 0, 0]
        # Scroll down to get all notes
        print('** search for all notes by {}...'.format(keyword))
        # input one search keyword
        page.ele('#search-input').clear()
        page.ele('#search-input').input('{}\n'.format(keyword))
        # page('').click(), if there is no '\n'
        print('*** trying to extract targeted info...')
        page.wait.load_start()
        notes_count = len(notes)-1        
        # to start
        notes_new = save_notes(keyword)
        notes_count += notes_new
        # loading all the notes              
        # while page.ele('@class=feeds-page').children()[-1].text.find('- THE END -') == -1:
        while notes_new != 0:
            print('** {} NEW note(s) recorded, {} notes saved in total'.format(notes_new, notes_count))
            print('** Please leave your webbrowser alone, still progressing...')
            if os.path.exists('pause.txt'):
                print('''\n\nDetected pause.txt, will pause and save the data
                    \nFor picking up next time, just rerun the script \n''')
                extract_notes()
                end()
            page.scroll.down(1000)
            time.sleep(random.uniform(1, 3))
            notes_new = save_notes(keyword)
            notes_count += notes_new            
        print('** {} NEW note(s) recorded, {} notes saved in total'.format(notes_new, notes_count))
        print('** ALL NOTES FOUND by', keyword)
        search_result_all[keyword][1] = search_result_all[keyword][0] + search_result_all[keyword][2]
        print('\n')
        for i in search_result_all.keys():            
            print(i, search_result_all[i])
        print('\n')
        extract_notes()

def note_contents_extract():
    # targeted data: content,  ip (if), post date, comments with author reply (if) - THE END
    print('** extract contents of the note')
    content_item = str(page.ele('@id=detail-desc').text).replace('\n', '。')
    ''' # may need to extract the tags at the bottom of notes
    try:
        tag_first = page.ele('@class=tag').text
    except:
        tag_first = ''
    if tag_first != '':
        tag_start = content_item.find(tag_first)'''
    # extract date and IP: '(编辑于) mm-dd','(编辑于) mm-dd ip', '(编辑于) x 天前 （ip）', '(编辑于) 昨天 hh:mm (ip)'; '(编辑于) 今天 hh:mm (ip)'
    dt=datetime.now() #datetime对象
    date_mix = str(page.ele('@class=date').text).split(' ')
    if '天前' in date_mix:
        m = date_mix.index('天前')
        date = dt.date() + timedelta(days=-int(date_mix[m-1]))
        if ('编辑于' in date_mix) and (len(date_mix) == 4):
            ip = date_mix[-1]
        elif ('编辑于' not in date_mix) and (len(date_mix) == 3):
            ip = date_mix[-1]
        else:
            ip = ''
    elif '昨天' in date_mix:
        date = date = dt.date() + timedelta(days=-1)
        if ('编辑于' in date_mix) and (len(date_mix) == 4):
            ip = date_mix[-1]
        elif ('编辑于' not in date_mix) and (len(date_mix) == 3):
            ip = date_mix[-1]
        else:
            ip = ''
    elif '今天' in date_mix:
        date = dt.date()
        if ('编辑于' in date_mix) and (len(date_mix) == 4):
            ip = date_mix[-1]
        elif ('编辑于' not in date_mix) and (len(date_mix) == 3):
            ip = date_mix[-1]
        else:
            ip = ''
    else:
        if '编辑于' in date_mix:
            date_temp = date_mix[1]
            if len(date_mix) == 3:
                ip = date_mix[-1]
            else:
                ip = ''
        else:
            date_temp = date_mix[0]
            if len(date_mix) == 2:
                ip = date_mix[-1]
            else:
                ip = ''
        date_temp2 = date_temp.split('-')
        if len(date_temp2) == 2:
            date = str(dt.year)+'-'+date_temp
        else:
            date = date_temp
    print('* post on', date, ip)
    return content_item, str(date), ip

def note_comments_extract(comments, ip): 
    # comments = [], ip
    def comments_load():
    # load all comments, scroll to current last comment and wait for loading
        i = 0
        while i != len(page.eles('@class=parent-comment')):
            if i != 0:            
                print('* loading comments, please leave the web browser alone...')
            i = len(page.eles('@class=parent-comment'))
            page.eles('@class=parent-comment')[-1].scroll.to_see()
            time.sleep(random.uniform(1, 15))

    if page.ele('@class=comments-el').text.find('点击评论') != -1:
        print('* no comment')
        ip = ip
    else:
        c_total = int(page.ele('@class=comments-container').ele('@class=total').text.split(' ')[1])
        print('* {} comment(s) in total, selecting thoses involved the author...'.format(c_total))
        # load all comments: to start
        comments_load()
        # loading all
        while page.ele('@class=comments-container').text.find('- THE END -') == -1:
            comments_load()
        # extract all comments
        comments_to_extract = page.eles('@class=parent-comment')        
        # each comment        
        for i in comments_to_extract:
            c_author = -1
            r_author = -1
            c_select = set()
            if len(i.children()) == 1:
            # comment with no reply                
                c_author = i.ele('@class=comment-inner-container').text.find('作者')
            else:
            # has both comment and reply                    
                c_author = i.ele('@class=comment-inner-container').text.find('作者')
                r_author = i.ele('@class=reply-container').text.find('作者')
            if c_author!= -1 and r_author != -1:
            # commented and replied by the author                
                comment_by_author = i.ele('@class=comment-inner-container').ele('@class=content').text
                reply_by_author = i.ele('@class=reply-container').ele('@class=content').text
                c_select = ('author: '+str(comment_by_author), 'author: '+str(reply_by_author))
            # commented by the author
            elif c_author != -1 and r_author == -1:
                comment_by_author = i.ele('@class=comment-inner-container').ele('@class=content').text
                c_select = ('author: '+str(comment_by_author))
            # replied by the author
            elif r_author != -1 and c_author == -1:
                ask = i.ele('@class=comment-inner-container').ele('@class=content').text
                reply_by_author = i.ele('@class=reply-container').ele('@class=content').text
                c_select = ('others: '+str(ask), 'author: '+str(reply_by_author))
            # check duplication           
            if len(c_select) != 0:
                comments.append(c_select)
                print('* {} comment(s) recorded'.format(len(comments)))
                # may have ip for replies
            if ip == '':
                if c_author != -1:
                    j = i.ele('@class=comment-inner-container').ele('@class=date').children()
                elif r_author != -1:
                    j = i.ele('@class=reply-container').ele('@class=date').children()
                else:
                    j = []
                if len(j) == 2:
                    try:
                        ip = i.ele('@class=location').text
                        if len(ip) != 0:
                            print('* detect IP address of author in replies', ip)
                    except:
                        pass                            
        if len(comments) == 0:
            print('* no comments involved the author')
    return ip

def note_open_and_save(note_id, note_info):
    # note_id = i : for i in page.eles('@class=cover ld mask')
    # note_info = notes_save[i]: ['author', 'author_profile', 'title', 'note_url - unaccessible']
    print('** try to record contents of the note and comments replied by the authors...')
    # open a note by clicking
    note_id.click()
    time.sleep(random.uniform(1, 2))

    # extract note info
    # save data temporarily
    comments_temp= []
    content_item, date, ip = note_contents_extract()
    if comments == 'Yes':
        print('** extract comments replied by the author...')
        ip = note_comments_extract(comments_temp, ip)
    search_date = str(datetime.today())
    note_contents.append((ip, date, note_info[0], note_info[1], note_info[2], note_info[3], content_item, comments_temp, search_date))   
    # close the note window
    time.sleep(random.uniform(1, 2))
    window_size = page.ele('@id=noteContainer').attr('style').find('translate(0px, 0px)')
    if  window_size != -1:
        # close the note tab for small windows mode
        page.ele('@class=close-box').click()    
    elif window_size == -1:
        # close the note for big windows mode
        page.ele('@class=close close-mask-dark').click()
    else:
        pass
    time.sleep(random.uniform(1, 2))

def search_notes_contents(notes_save):
    def note_found():
        print('*** found', note_title)
        note_url = '/'+note_url_full.lstrip('https://www.xiaohongshu.com')
        note_info = [author, author_profile, note_title, note_url_full]
        note_id = page.ele('@href='+note_url).parent()
        note_id.scroll.to_see()
        note_open_and_save(note_id, note_info)
        # for each note targeted, need to reload the profile page because the main page will be scrolled when loading comments

    # Get all notes
    notes_not_found = 0
    for n in range(len(notes_save)): # each note
        try:          
            note_loading = n
            i = notes_save[n]            
            if i[0] != 'author': # head
                author = i[0]
                author_profile = i[1]
                note_title = i[2]
                note_url_full = i[3]
                print('\n** open the profile of', author)
                # open author profile to start
                page.get(author_profile)
                # find the note
                print('\n** looking for notes', note_title)                           
                # loading profile to search and save the note
                j = 'loading' 
                while j == 'loading':
                    # check
                    notes_loaded = page.eles('tag:a').get.links()
                    if note_url_full in notes_loaded:
                        note_found()
                        j = 'note found'
                    else:
                        print('**** loading notes, please leave the web browser alone...')
                        k1 = page.eles('@class=note-item')[-1]
                        k1.scroll.to_see()
                        time.sleep(random.uniform(1, 2))                            
                        # check all loaded
                        k2 = page.eles('@class=note-item')[-1]
                        if k2 == k1:
                            # last check
                            notes_loaded = page.eles('tag:a').get.links()
                            if note_url_full in notes_loaded:
                                note_found()
                                j = 'note found'
                            else:
                                j = 'note lost'
                        else:
                            pass # check in the next round
                
                # can't find the note
                if j == 'note lost':
                    print('\n** the note may be deleted')
                    notes_not_found += 1
                    nd = 'not found'
                    search_date = str(datetime.today())
                    note_contents.append((nd, nd, author, author_profile, note_title, note_url_full, nd, nd, search_date)) 
                
                print('** {} notes recorded, {} note(s) left'.format(len(note_contents)-1, len(notes_save)-len(note_contents)))
                human()

                if os.path.exists('pause.txt'):
                    print('''\n\nDetected pause.txt, will pause and save the data
                        \nFor picking up next time, just rerun the script \n''')
                    notes_save2 = [('author', 'author_profile', 'title', 'note_url - unaccessible')]
                    for n in range(note_loading, len(notes_save)):
                        notes_save2.append(notes_save[n])
                    excel_write(search_result_left, notes_save2)
                    end()               
        except:
            print('\n* something went wrong with note {}, will skip.'.format(notes_save[note_loading][2]))
            notes_wrong.append(notes_save[note_loading])
            note_loading += 1
            if (note_loading+1) != len(notes_save):
                notes_save2 = [('author', 'author_profile', 'title', 'note_url - unaccessible')]
                for n in range(note_loading, len(notes_save)):
                    notes_save2.append(notes_save[n])
                search_notes_contents(notes_save2)
            if os.path.exists('pause.txt'):
                print('''\n\nDetected pause.txt, will pause and save the data
                    \nFor picking up next time, just rerun the script \n''')
                excel_write(search_result_left, notes_save2)            
    # save results
    print('\n\n* all note targeted are searched, {} saved, {} not found'.format(len(notes_save)-1-notes_not_found, notes_not_found))
    excel_write(notes_result, note_contents)
    if len(notes_wrong) > 1:
        excel_write(search_wrong, notes_wrong)
    end()

def input_wait(remind, defaut):
    """input or leap over"""
    # CODE from: https://blog.csdn.net/weixin_39858881/article/details/107152961
    import func_timeout

    @func_timeout.func_set_timeout(10)
    def askChoice():
        return input(remind)

    try:
        s = askChoice()
    except func_timeout.exceptions.FunctionTimedOut as e:
        s = defaut
    print('\n', s)
    return s

def update_check(user_name, repository_name, file_path):  
    # Gitee API URL to get latest commit
    '''user_name = 'sidchen0122'
    repository_name = 'red-note-spider'
    file_path = 'RedNoteSpider.exe'''    
    print(f'\n* checking updates for {file_path} ...')
    url = f'https://gitee.com/api/v5/repos/{user_name}/{repository_name}/commits?path={file_path}'
    print(url)

    try:
        response = requests.get(url)
        if response.status_code == 200:
            latest_commit = response.json()[0]["commit"]["author"]["date"].split("T")[0]  # Get commit hash
            print(f"Latest update: {latest_commit}")
            version_latest = datetime.strptime(latest_commit, '%Y-%m-%d')
            version_pre = datetime.strptime(progress[file_path], '%Y-%m-%d')
            if version_latest != version_pre:                
                download_url = f"https://gitee.com/{user_name}/{repository_name}/raw/master/{file_path}"
                print(f'New version found! raising download from {download_url} ...')
                webbrowser.open(download_url)
        else:
            print("Failed to check for updates:", response.text)
    except Exception as e:
        print(f"Error checking updates: {e}")


if __name__ == "__main__":
    print('* 本脚本用于获取小红书笔记数据')
    # check and download the updated progress
    progress = {'RedNoteSpider.exe': '2025-04-05', 'README.md': '2025-04-05'}
    print('Project address: https://gitee.com/sidchen0122/red-note-spider.git')
    for i in progress:
        print(f'* {i}，版本号{progress[i]}，')
    print('*** 有限技术支持：sidchen0 @ qq.com ***')
    print('\nIf you wanna pause the script when searching for note contents, please create a pause.txt file in the same folder')
    for i in progress:
        update_check('sidchen0122', 'red-note-spider', i)
    
    # Change current working directory to the script's directory
    # os.chdir(os.path.dirname(os.path.abspath(__file__)))

    # Determine the base directory of the executable or script
    if getattr(sys, 'frozen', False):  # If running as a bundled executable
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    # Change the working directory to the base directory
    os.chdir(base_dir)


    # to save data0
    note_contents = [('ip', 'post_date','author', 'author_profile', 'title', 'note_url - unaccessible', 'content', 'comments_selected', 'search_date')]
    notes_wrong = [('author', 'author_profile', 'title', 'note_url - unaccessible')]
    # for function human()
    start = datetime.now()
    nap = {'start': start, 'nap': start, 'sleep': start}

    search_result = 'data/search_result.xlsx'
    search_result_left = 'data/search_result_left.xlsx'
    search_wrong = 'data/notes_save_wrong.xlsx'
    notes_result = 'data/notes_contents and comments.xlsx'

    # second round
    if os.path.exists(search_result_left):
        print('\n* Do you want to pick up data collection and get all notes\' contents?')
        mode_choices = ['pick up data collection', 'No, start a new round']
        mode_chosen = human_choose(mode_choices, 1)
        print('* mode chosen:', mode_choices[mode_chosen[0]])
        if mode_chosen[0] == 0:
            mode = 'pick up now'
        else:
            mode = 'normal'
        # change file names to continue
        if os.path.exists(search_result):
            os.rename(search_result, search_result.strip('.xlsx')+'_pre.xlsx')
        os.rename(search_result_left, search_result)
    elif os.path.exists(search_result) == True and os.path.exists(notes_result) == False:
        print('\n* Do you want to pick up data collection and get all notes\' contents?')
        mode_choices = ['pick up data collection', 'No, start a new round']
        mode_chosen = human_choose(mode_choices, 1)
        print('* mode chosen:', mode_choices[mode_chosen[0]])
        if mode_chosen[0] == 0:
            mode = 'pick up now'
        else:
            mode = 'normal'
    else:
        mode = 'normal'
    
    if os.path.exists(notes_result):
        os.rename(notes_result, notes_result.strip('.xlsx')+'_pre.xlsx')
    if os.path.exists(search_wrong):
        os.rename(search_wrong, search_wrong.strip('.xlsx')+'_pre.xlsx')

    # first round
    if mode == 'normal':
        print ('\n* For this round, do you want to save all notes\' contents after searching by the keywords\nor pick up later by running the script again?')
        mode_choices = ['save all notes\' contents after searching by the keywords', 'pick up later by run the script again']
        mode_chosen = human_choose(mode_choices, 1)
        print('* mode chosen:', mode_choices[mode_chosen[0]])
        if mode_chosen[0] == 0:
            mode = 'normal'
        else:
            mode = 'pick up later'

    if mode != 'pick up later':
        print('\nWhen save the details of notes, do you want to save the comments replied by the author?')
        mode_choices = ['with comments (may take longer)', 'without comments']
        mode_chosen = human_choose(mode_choices, 1)
        print('* chosen:', mode_choices[mode_chosen[0]])
        if mode_chosen[0] == 0:
            comments = 'Yes'
        else:
            comments = 'No'

    print('\n* try to start the Chrome webbrowser...')
    '''path = r'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe'  # 请改为你电脑内Chrome可执行文件路径
    co = ChromiumOptions().set_browser_path(path)'''
    # co = SessionOptions().set_browser_path(path).save() change the default browser permanently
    try:
        page = WebPage()
    except:
        input('* cannot start the Chrome browser properly, please check')
        exit()
    website_targeted = 'https://www.xiaohongshu.com/explore'
    print('\n* open', website_targeted)
    page.get(website_targeted)
    # log in check
    log_in_check()

    if mode == 'pick up later':
        keywords = keywords_load()
        search_keywords(keywords)
        end()
    elif mode == 'normal':
        keywords = keywords_load()
        search_keywords(keywords)
        # pick up data spider
        notes_save = []
        excel_read(search_result, ['author', 'author_profile', 'title', 'note_url - unaccessible'], notes_save)
        search_notes_contents(notes_save)
    elif mode == 'pick up now':
        # pick up data spider
        notes_save = []
        excel_read(search_result, ['author', 'author_profile', 'title', 'note_url - unaccessible'], notes_save)
        search_notes_contents(notes_save)