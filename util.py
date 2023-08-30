import requests
import bs4
import xlsxwriter as excel

class Util:
    __mwe_player_count = 0
    def downloadFile(self, file_url: str, file_type: int, name: str) -> str: # 0 for image, 1 for sound, 2 for video
        file = requests.get(file_url, stream=True)
        saved_name = name + '.' + ('png' if file_type == 0 else 'ogg' if file_type == 1 else 'ogv')
        with open("media/" + saved_name, 'wb') as f:
            for chunk in file.iter_content(chunk_size=1024):
                f.write(chunk)
        return saved_name

    def KhoiDong(self, cont: bs4.element.Tag) -> None:
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
            tr_child = tr.find_all('th') # check if this tr tag is round number or question content
            if len(tr_child) == 1: # round number
                for sh in sheet:
                    sh.merge_range(row_count, 0, row_count, 1, tr_child[0].text)
            else: # question content
                tr_child = tr.find_all('td') # questions content stored in td tags
                check_image = tr_child[0].find_all('img') # check image
                if len(check_image) != 0: # question contains image
                    img_src = check_image[0].attrs['data-src']
                    img_name = check_image[0].attrs['alt']
                    img_name = self.downloadFile(img_src, 0, img_name)
                    for sh in sheet:
                        sh.write_url(row_count, 0, "media/" + img_name, string=tr_child[0].text)
                    ans_sheet.write_string(row_count, 1, tr_child[1].text)
                else:
                    check_sound = tr_child[0].find_all('div') # check sound
                    if len(check_sound) != 0: # question contains sound
                        sound_id = "mwe_player_" + str(self.__mwe_player_count)
                        self.__mwe_player_count += 1
                        sound_id = tr_child[0].find(id=sound_id)
                        sound_id = sound_id.children
                        sound_id = next(sound_id).attrs['src']
                        sound_name = self.downloadFile(sound_id, 1, "KhoiDong-" + str(self.__mwe_player_count))
                        for sh in sheet:
                            sh.write_url(row_count, 0, "media/" + sound_name, string=tr_child[0].text)
                        ans_sheet.write_string(row_count, 1, tr_child[1].text)
                    else: # normal question - just text
                        for sh in sheet:
                            sh.write_string(row_count, 0, tr_child[0].text)
                        ans_sheet.write_string(row_count, 1, tr_child[1].text)
            row_count += 1
        workbook.close()

    def VCNV(self, cont: bs4.element.Tag) -> None:
        workbook = excel.Workbook('VCNV.xlsx')
        quest_sheet = workbook.add_worksheet('Quest')
        ans_sheet = workbook.add_worksheet('Ans')
        sheet = [quest_sheet, ans_sheet]
        row_count = 0
        tr_tags = (cont.find_all('tr'))[4:] # remove unneccessary four first tags
        # adding keyword info and basics
        num_of_letter = tr_tags[-2].find_all('img')[0].attrs['alt'][0] # extract number of letters
        # get final image and download
        image_src = tr_tags[-3].find_all('img')[0].attrs['data-src']
        image_name = self.downloadFile(image_src, 0, "HinhCNV")
        answer = tr_tags[-1].text
        for sh in sheet:
            sh.merge_range(row_count, 0, row_count, 2, "CNV có " + num_of_letter)
            sh.write_string(row_count + 2, 0, 'Hàng ngang')
            sh.write_string(row_count + 2, 1, 'Câu hỏi')
            sh.write_string(row_count + 2, 2, 'Đáp án')
        quest_sheet.merge_range(row_count + 1, 0, row_count + 1, 2, "")
        ans_sheet.merge_range(row_count + 1, 0, row_count + 1, 2, answer)
        row_count += 3

        # done with result and final image, can remove last rows
        tr_tags = tr_tags[:-4]
 
        # adding quests and answers
        for tr in tr_tags:
            td_tags = tr.find_all('td')
            no_let = td_tags[0].text
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
                file_name = self.downloadFile(sound, 1, "VCNV_" + str(self.__mwe_player_count))
                for sh in sheet:
                    sh.write_url(row_count, 1, "media/" + file_name, string=td_tags[1].text)
            else:
                for sh in sheet:
                    sh.write_string(row_count, 1, td_tags[1].text)
            row_count += 1
        ans_sheet.write_url(row_count, 1, "media/" + image_name, string='Hình ảnh CNV')
        workbook.close()

    def TangToc(self, cont: bs4.element.Tag, year: int) -> None:
        workbook = excel.Workbook('TangToc.xlsx')
        quest_sheet = workbook.add_worksheet('Quest')
        ans_sheet = workbook.add_worksheet('Ans')
        sheet = [quest_sheet, ans_sheet]
        for sh in sheet:
            sh.write_string(0, 0, "Câu hỏi")
            sh.write_string(0, 1, "Đáp án")
        tr_tags = cont.find_all('tr')[2:] # remove unneccessary 2 first tags
        for i in range(1, 5):
            tr = tr_tags[i - 1]
            td_tags = tr.find_all('td')
            img = td_tags[0].find_all('img')
            if len(img) > 0: # image question
                img_src = img[0].attrs['data-src']
                image_name = self.downloadFile(img_src, 0, "Tangtoc" + str(i))
                for sh in sheet:
                    sh.write_url(i, 0, "media/" + image_name, string=("TangToc" + str(i)))
                answer_text = td_tags[1].text
                check_ans = td_tags[1].find_all('img')
                if len(check_ans) == 0:
                    ans_sheet.write_string(i, 1, answer_text)
                else:
                    ans_src = check_ans[0].attrs['data-src']
                    answer_img = self.downloadFile(ans_src, 0, "Tangtoc" + str(i) + "Dapan")
                    ans_sheet.write_url(i, 1, "media/" + answer_img, string=((answer_text if answer_text != "" else "") + "DapAnTangToc" + str(i)))
            else: # video question
                vid_id = "mwe_player_" + str(self.__mwe_player_count)
                self.__mwe_player_count += 1
                vid_src = tr.find_all('video')[0].attrs['src']
                vid_name = self.downloadFile(vid_src, 2, "Tangtoc" + str(i))
                for sh in sheet:
                    sh.write_url(i, 0, "media/" + vid_name, string=("TangToc" + str(i)))
                answer = td_tags[1].text
                ans_sheet.write_string(i, 1, answer)
        workbook.close()

    def VeDich(self, cont: bs4.element.Tag) -> None:
        workbook = excel.Workbook('VeDich.xlsx')
        quest_sheet = workbook.add_worksheet('Quest')
        ans_sheet = workbook.add_worksheet('Ans')
        sheet = [quest_sheet, ans_sheet]
        for sh in sheet:
            sh.write_string(0, 0, "Số điểm")
            sh.write_string(0, 1, "Câu hỏi")
            sh.write_string(0, 2, "Đáp án")
        tr_tags = cont.find_all('tr')[2:] # remove unneccessary two first tags
        no_main = len(tr_tags)
        for row_idx in range(0, no_main):
            if row_idx % (no_main / 4) == 0: # players' name and equivalent points selected
                player_name = tr_tags[row_idx].text.split('(')[0]
                quest_sheet.merge_range(row_idx + 1, 0, row_idx + 1, 2, player_name)
                ans_sheet.merge_range(row_idx + 1, 0, row_idx + 1, 2, player_name)
                point_img = tr_tags[row_idx].find_all('img')
                for quest_no in range (0, len(point_img)):
                    point = point_img[quest_no].attrs['alt'].split(' ')[4]
                    for sh in sheet:
                        sh.write_string(row_idx + quest_no + 2, 0, point + " điểm")
            else: # quests
                td_tags = tr_tags[row_idx].find_all('td')
                # handle question content
                check_media_quest = td_tags[0].find_all('video')
                if (len(check_media_quest) != 0):
                    id_str = "mwe_player_" + str(self.__mwe_player_count)
                    self.__mwe_player_count += 1
                    vid_src = check_media_quest[0].attrs['src']
                    vid_name = self.downloadFile(vid_src, 2, "CauHoi" + str(int(row_idx / 4) + 1) + "-" + str(row_idx % 4))
                    for sh in sheet:
                        sh.write_url(row_idx + 1, 1, "media/" + vid_name, string=(td_tags[0].text.split('.')[0]))
                else:
                    for sh in sheet:
                        sh.write_string(row_idx + 1, 1, td_tags[0].text)

                # handle answer content
                check_media_ans = td_tags[1].find_all('video')
                if (len(check_media_ans) != 0):
                    id_str = "mwe_player_" + str(self.__mwe_player_count)
                    self.__mwe_player_count += 1
                    vid_src = check_media_ans[0].attrs['src']
                    vid_name = self.downloadFile(vid_src, 2, "Dapan" + str(int(row_idx / 4) + 1) + "-" + str(row_idx % 4))
                    ans_sheet.write_url(row_idx + 1, 2, "media/" + vid_name, string=(vid_name))
                else:
                    ans_sheet.write_string(row_idx + 1, 2, td_tags[1].text)
        workbook.close()

    def CauHoiPhu(self, cont: bs4.element.Tag) -> None:
        workbook = excel.Workbook('CauHoiPhu.xlsx')
        quest_sheet = workbook.add_worksheet('Quest')
        ans_sheet = workbook.add_worksheet('Ans')
        sheet = [quest_sheet, ans_sheet]
        for sh in sheet:
            sh.write_string(0, 0, 'Câu hỏi')
            sh.write_string(0, 1, 'Đáp án')
        row_count = 1
        tr_tags = (cont.find_all('tr'))[3:] # remove unneccessary three first tags
        for tr in tr_tags:
            td_tags = tr.find_all('td')
            check_image = td_tags[0].find_all('img') # check image
            if len(check_image) != 0:
                img_src = check_image[0].attrs['data-src']
                img_name = check_image[0].attrs['alt']
                img_name = self.downloadFile(img_src, 0, img_name)
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
                    sound_name = self.downloadFile(sound_id, 1, "KhoiDong-" + str(self.__mwe_player_count))
                    for sh in sheet:
                        sh.write_url(row_count, 0, "media/" + sound_name, string=td_tags[0].text)
                    ans_sheet.write_string(row_count, 1, td_tags[1].text)
                else:
                    for sh in sheet:
                        sh.write_string(row_count, 0, td_tags[0].text)
                    ans_sheet.write_string(row_count, 1, td_tags[1].text)
            row_count += 1
        workbook.close()

if __name__ == '__main__':
    print("For imported only!")