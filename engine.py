import bs4
import requests
import os
import engine_21 as o21
import engine_22 as o22

class Engine():
    def __init__(self) -> None:
        raise Exception("Wrong parameters!")

    def __init__(self, year: int, quarter: int = 0, month: int = 0, week: int = 0, up_to: bool = False) -> None:
        self.__year = year
        self.__quarter = quarter
        self.__month = month
        self.__week = week
        if not self.__checkInput(): raise Exception("Invalid input!")
        self.__const_root = os.getcwd()
        self.__root = os.getcwd()
        print("Constructor: ", self.__root)
        self.__initDirSystem()
        if up_to == True:
            if self.__quarter == 0: 
                self.__run()
                os.chdir(self.__const_root)
                self.__root = self.__const_root
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
                            if self.__checkInput(): 
                                self.__run()
                                os.chdir(self.__const_root)
                                self.__root = self.__const_root
        else:
            self.__run()
            os.chdir(self.__const_root)
            self.__root = self.__const_root

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

    def __run(self) -> None:
        # target dir in directory system
        target_dir = ""
        if self.__quarter == 0:
            target_dir = "CKN"
        else: 
            match = (self.__week, self.__month, self.__quarter)
            for m in match:
                target_dir += ("x" if m == 0 else str(m))
        print("Run: ", self.__root)
        os.chdir(self.__root + "/O" + str(self.__year))
        self.__root = os.getcwd()
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
        is_final = False
        self.__url = "https://duong-len-dinh-olympia.fandom.com/vi/wiki/Olympia_" + str(self.__year) + "/"
        if self.__quarter == 0: 
            self.__url += "Chung_kết"
            is_final = True
        else: 
            if self.__week > 0: self.__url += "Tuần_" + str(self.__week) + "_"
            if self.__month > 0: self.__url += "Tháng_" + str(self.__month) + "_"
            if self.__quarter > 0: self.__url += "Quý_" + str(self.__quarter)
        #
        page = requests.get(self.__url)
        page_html = bs4.BeautifulSoup(page.content, 'lxml')
        content = page_html.find_all('table', class_='sectiontable')
        if len(content) == 0:
            print("Match " + target_dir + " not shown yet!")
            return
        
        # need modifying for final match:
        if is_final:
            content = content[1:4] + content[5:] # remove livespots and quests for livespots
        
        try: os.mkdir('media')
        except FileExistsError:
            pass

        engine_21 = o21.Engine_21()
        engine_22 = o22.Engine_22()

        content_index = 1 # round number: 1 - KĐ, 2 - VCNV, 3 - TT, 4 - VĐ, 5 - CHP
        for cont in content:
            if content_index == 1:
                engine_21.KhoiDong(cont) if self.__year == 21 else engine_22.KhoiDong(cont)
            elif content_index == 2:
                engine_21.VCNV(cont) if self.__year == 21 else engine_22.VCNV(cont)
            elif content_index == 3:
                engine_21.TangToc(cont, self.__year) if self.__year == 21 else engine_22.TangToc(cont, self.__year)
            elif content_index == 4:
                engine_21.VeDich(cont) if self.__year == 21 else engine_22.VeDich(cont)
            elif content_index == 5:
                engine_21.CauHoiPhu(cont) if self.__year == 21 else engine_22.CauHoiPhu(cont)
            content_index += 1
        print(target_dir + " Done")

if __name__ == '__main__':
    print("For imported only!")