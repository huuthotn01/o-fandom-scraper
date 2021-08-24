import bs4
import requests
import xlsxwriter as excel
import os

class Engine:
    def __init__(self) -> None:
        raise Exception("Wrong parameters!")

    def __init__(self, year: int, quarter: int = 0, month: int = 0, week: int = 0, up_to: bool = False) -> None:
        self.__year = year
        self.__quarter = quarter
        self.__month = month
        self.__week = week
        if not self.__checkInput(): raise Exception("Invalid input!")
        self.__root = os.getcwd()
        self.__initDirSystem()
        self.__mwe_player_count = 0
        if up_to == True:
            if self.__quarter == 0: self.__run()
            match = (self.__quarter, self.__month, self.__week)
            match_code = 0
            for i in range(3): match_code += (10**i) * (5 if match[2 - i] == 0 else match[2 - i])
            for qu in range(1, 5):
                for mo in range(4):
                    for we in range(4):
                        self.__quarter = qu
                        self.__month = mo
                        self.__week = we
                        curr_match = (100 * qu) + 10*(5 if mo == 0 else mo) + (5 if we == 0 else we)
                        if curr_match <= match_code:
                            if self.__checkInput(): self.__run()
        else:
            self.__run()

    def __checkInput(self) -> bool:
        if self.__year <= 0 | (self.__quarter not in (0, 1, 2, 3, 4)) | (self.__month not in (0, 1, 2, 3)) | (self.__week not in (0, 1, 2, 3)): return False
        if self.__quarter == 0:
            if (self.__month**2 + self.__week**2) != 0: return False
        else:
            if (self.__month == 0) & (self.__week != 0): return False
        return True

    def __initDirSystem(self) -> None:
        quarter = (1, 2, 3, 4)
        week_month = (1, 2, 3)
        path = "O" + str(self.__year)
        final_folder = "CKN"
        final_folder = os.path.join(path, final_folder)
        os.makedirs(final_folder, exist_ok=True)
        for q in quarter:
            qu_dir = "Quý " + str(q)
            qu_dir = os.path.join(path, qu_dir)
            for m in week_month:
                mo_dir = "Tháng " + str(m)
                mo_dir = os.path.join(qu_dir, mo_dir)
                for w in week_month:
                    we_dir = str(w) + str(m) + str(q)
                    we_dir = os.path.join(mo_dir, we_dir)
                    os.makedirs(we_dir, exist_ok=True)
                mo_code = "x" + str(m) + str(q)
                mo_path = os.path.join(mo_dir, mo_code)
                os.makedirs(mo_path, exist_ok=True)
            qu_code = "xx" + str(q)
            qu_path = os.path.join(qu_dir, qu_code)
            os.makedirs(qu_path, exist_ok=True)
        print("Initialized directories!")

    def __downloadFile(self, file_url: str, file_type: int, name: str) -> str: # 0 for image, 1 for sound, 2 for video
        file = requests.get(file_url, stream=True)
        saved_name = name + '.' + ('png' if file_type == 0 else 'ogg' if file_type == 1 else 'ogv')
        with open("media/" + saved_name, 'wb') as f:
            for chunk in file.iter_content(chunk_size=1024):
                f.write(chunk)
        return saved_name

    def __run(self) -> None:
        # target dir in directory system
        target_dir = ""
        if self.__quarter == 0:
            target_dir = "CKN"
        else: 
            match = (self.__week, self.__month, self.__quarter)
            for m in match:
                target_dir += ("x" if m == 0 else str(m))
        for dirpath, dirnames, filenames in os.walk(self.__root):
            done = False
            for dirname in dirnames:
                if dirname == target_dir:
                    done = True
                    dirname = os.path.join(dirpath, dirname)
                    os.chdir(dirname)
                    break
            if done: break
        # target URL
        self.__url = "https://duong-len-dinh-olympia.fandom.com/vi/wiki/Năm_thứ_" + str(self.__year) + "/"
        if self.__week > 0: self.__url += "Tuần_" + str(self.__week) + "_"
        if self.__month > 0: self.__url += "Tháng_" + str(self.__month) + "_"
        if self.__quarter > 0: self.__url += "Quý_" + str(self.__quarter)
        #
        page = requests.get(self.__url)
        page_html = bs4.BeautifulSoup(page.content, 'lxml')
        content = page_html.find_all('table', class_='wikitable')
        if len(content) == 4:
            print("Match " + target_dir + " not shown yet!")
            return
        content_index = 0
        try: os.mkdir('media')
        except FileExistsError:
            pass
        for cont in content:
            if content_index == 1:
                self.__KhoiDong(cont)
            elif content_index == 2:
                self.__VCNV(cont)
            elif content_index == 3:
                self.__TangToc(cont)
            elif content_index == 4:
                self.__VeDich(cont)
            content_index += 1
        self.__mwe_player_count = 0
        print(target_dir + " Done")

    def __KhoiDong(self, cont: bs4.element.Tag) -> None:
        workbook = excel.Workbook('KhoiDong.xlsx')
        quest_sheet = workbook.add_worksheet('Quest')
        ans_sheet = workbook.add_worksheet('Ans')
        sheet = [quest_sheet, ans_sheet]
        for sh in sheet:
            sh.write_string(0, 0, 'Câu hỏi')
            sh.write_string(0, 1, 'Đáp án')
        row_count = 1
        tr_tags = (cont.find_all('tr'))[2:] # remove unneccessary two first tags
        for tr in tr_tags:
            td_tags = tr.find_all('td')
            if len(td_tags) == 1:
                for sh in sheet:
                    sh.merge_range(row_count, 0, row_count, 1, td_tags[0].text)
            else:
                check_image = td_tags[0].find_all('img') # check image
                if len(check_image) != 0:
                    img_src = check_image[0].attrs['data-src']
                    img_name = check_image[0].attrs['alt']
                    img_name = self.__downloadFile(img_src, 0, img_name)
                    for sh in sheet:
                        sh.write_url(row_count, 0, "media/" + img_name, string=td_tags[0].text)
                    ans_sheet.write_string(row_count, 1, td_tags[1].text)
                else:
                    check_sound = td_tags[0].find_all('div') # check sound
                    if len(check_sound) != 0:
                        sound_id = "mwe_player_" + str(self.__mwe_player_count)
                        self.__mwe_player_count += 1
                        sound_id = td_tags[0].find(id=sound_id)
                        sound_id = sound_id.children
                        sound_id = next(sound_id).attrs['src']
                        sound_name = self.__downloadFile(sound_id, 1, "KhoiDong-" + str(self.__mwe_player_count))
                        for sh in sheet:
                            sh.write_url(row_count, 0, "media/" + sound_name, string=td_tags[0].text)
                        ans_sheet.write_string(row_count, 1, td_tags[1].text)
                    else:
                        for sh in sheet:
                            sh.write_string(row_count, 0, td_tags[0].text)
                        ans_sheet.write_string(row_count, 1, td_tags[1].text)
            row_count += 1
        workbook.close()
    
    def __VCNV(self, cont: bs4.element.Tag) -> None:
        workbook = excel.Workbook('VCNV.xlsx')
        quest_sheet = workbook.add_worksheet('Quest')
        ans_sheet = workbook.add_worksheet('Ans')
        sheet = [quest_sheet, ans_sheet]
        row_count = 0
        tr_tags = (cont.find_all('tr'))[4:] # remove unneccessary four first tags
        # adding keyword info and basics
        num_of_letter = (tr_tags[-2].text).split('(')[1].split(')')[0]
        answer = tr_tags[-1].text
        for sh in sheet:
            sh.merge_range(row_count, 0, row_count, 2, "CNV có " + num_of_letter)
            sh.write_string(row_count + 2, 0, 'Hàng ngang')
            sh.write_string(row_count + 2, 1, 'Câu hỏi')
            sh.write_string(row_count + 2, 2, 'Đáp án')
        quest_sheet.merge_range(row_count + 1, 0, row_count + 1, 2, "")
        ans_sheet.merge_range(row_count + 1, 0, row_count + 1, 2, answer)
        row_count += 3
        # adding quests and answers
        for tr in range(5):
            if (len(tr_tags) == 10) & (tr == 4): break
            td_tags = tr_tags[tr].find_all('td')
            if tr == 4:
                no_let = td_tags[0].text
                for sh in sheet:
                    sh.write_string(row_count, 0, "Ô trung tâm: " + no_let + " chữ cái")
            else:
                no_let = td_tags[0].text.split('(')[1].split(')')[0]
                for sh in sheet:
                    sh.write_string(row_count, 0, no_let)
            ans_sheet.write_string(row_count, 2, td_tags[2].text)
            check_media = td_tags[1].find_all('div')
            if len(check_media) != 0:
                id_str = "mwe_player_" + str(self.__mwe_player_count)
                self.__mwe_player_count += 1
                sound = td_tags[1].find(id=id_str)
                sound = sound.children
                sound = next(sound).attrs['src']
                file_name = self.__downloadFile(sound, 1, "VCNV_" + str(self.__mwe_player_count))
                for sh in sheet:
                    sh.write_url(row_count, 1, "media/" + file_name, string=td_tags[1].text)
            else:
                for sh in sheet:
                    sh.write_string(row_count, 1, td_tags[1].text)
            row_count += 1
        image_src = tr_tags[-3].find_all('img')[0].attrs['data-src']
        image_name = self.__downloadFile(image_src, 0, "HinhCNV")
        ans_sheet.write_url(row_count, 1, "media/" + image_name, string='Hình ảnh CNV')
        workbook.close()

    def __TangToc(self, cont: bs4.element.Tag) -> None:
        workbook = excel.Workbook('TangToc.xlsx')
        quest_sheet = workbook.add_worksheet('Quest')
        ans_sheet = workbook.add_worksheet('Ans')
        sheet = [quest_sheet, ans_sheet]
        for sh in sheet:
            sh.write_string(0, 0, "Câu hỏi")
            sh.write_string(0, 1, "Đáp án")
        tr_tags = cont.find_all('tr')[2:] # remove unneccessary 2 first tags
        for i in range(1, 5):
            if i == 1 or i == 3:
                img_src = (tr_tags[i - 1].find_all('td')[0]).find('img').attrs['data-src']
                image_name = self.__downloadFile(img_src, 0, "Tangtoc" + str(i))
                for sh in sheet:
                    sh.write_url(i, 0, "media/" + image_name, string=("TangToc" + str(i)))
                answer = tr_tags[i - 1].find_all('td')[1]
                check_ans = answer.find_all('img')
                if not check_ans:
                    ans_sheet.write_string(i, 1, answer.text)
                else:
                    ans_src = check_ans[0].attrs['data-src']
                    answer_img = self.__downloadFile(ans_src, 0, "Tangtoc" + str(i) + "Dapan")
                    ans_sheet.write_url(i, 1, "media/" + answer_img, string=("DapAnTangToc" + str(i)))
            else:
                vid_id = "mwe_player_" + str(self.__mwe_player_count)
                self.__mwe_player_count += 1
                vid_src = tr_tags[i - 1].find(id=vid_id)
                source = vid_src.children
                vid_src = next(source).attrs['src']
                vid_name = self.__downloadFile(vid_src, 2, "Tangtoc" + str(i))
                for sh in sheet:
                    sh.write_url(i, 0, "media/" + vid_name, string=("TangToc" + str(i)))
                answer = tr_tags[i - 1].find_all('td')[1].text
                ans_sheet.write_string(i, 1, answer)
        workbook.close()

    def __VeDich(self, cont: bs4.element.Tag) -> None:
        workbook = excel.Workbook('VeDich.xlsx')
        quest_sheet = workbook.add_worksheet('Quest')
        ans_sheet = workbook.add_worksheet('Ans')
        sheet = [quest_sheet, ans_sheet]
        for sh in sheet:
            sh.write_string(0, 0, "Số điểm")
            sh.write_string(0, 1, "Câu hỏi")
            sh.write_string(0, 2, "Đáp án")
        tr_tags = cont.find_all('tr')[2:] # remove unneccessary two first tags
        for row_idx in range(0, 16):
            if row_idx % 4 == 0: # players' name and equivalent points selected
                player_name = tr_tags[row_idx].text.split('(')[0]
                quest_sheet.merge_range(row_idx + 1, 0, row_idx + 1, 2, player_name)
                ans_sheet.merge_range(row_idx + 1, 0, row_idx + 1, 2, player_name)
                point_img = tr_tags[row_idx].find_all('img')
                for quest_no in range (0, 3):
                    point = point_img[quest_no].attrs['alt'].split(' ')[1]
                    for sh in sheet:
                        sh.write_string(row_idx + quest_no + 2, 0, point + " điểm")
            else: # quests
                td_tags = tr_tags[row_idx].find_all('td')
                check_media = td_tags[0].find_all('div')
                if (len(check_media) != 0):
                    id_str = "mwe_player_" + str(self.__mwe_player_count)
                    self.__mwe_player_count += 1
                    vid_src = td_tags[0].find(id=id_str)
                    source = vid_src.children
                    vid_src = next(source).attrs['src']
                    vid_name = self.__downloadFile(vid_src, 2, "CauHoi" + str(int(row_idx / 4) + 1) + "-" + str(row_idx % 4))
                    for sh in sheet:
                        sh.write_url(row_idx + 1, 1, "media/" + vid_name, string=td_tags[0].text)
                else:
                    for sh in sheet:
                        sh.write_string(row_idx + 1, 1, td_tags[0].text)
                ans_sheet.write_string(row_idx + 1, 2, td_tags[1].text)
        workbook.close()

if __name__ == '__main__':
    print("For imported only!")
    exit()