# web driver imports
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
# excel file imports
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
# time import
from time import sleep
#test
from selenium.webdriver.common.by import By

if __name__ == '__main__':
    #create a config.json, read that json, and pass in data to here
    #import json file
    import json 
    with open ('config.json') as json_file:
        data = json.load(json_file)
        gmail_id = data['gmail_id']
        gmail_password = data['gmail_password']
        meet_link = data['meet_link']

    # Some necessary things for automation with google driver
    opt = Options()
    opt.add_argument("--disable-infobars")
    opt.add_argument("start-maximized")
    opt.add_argument("--disable-extensions")

    # This part allows the notifications, mic and camera permissions
    # Pass the argument 1 to allow and 2 to block
    opt.add_experimental_option("prefs", {
        "profile.default_content_setting_values.media_stream_mic": 1,
        "profile.default_content_setting_values.media_stream_camera": 1,
        "profile.default_content_setting_values.geolocation": 1,
        "profile.default_content_setting_values.notifications": 1
      })

    # Sign in to google
    driver = webdriver.Chrome(options=opt, executable_path=r"C:\Users\Loe Hui Ying\Downloads\chromedriver.exe") # eg. r"C:\Users\hp\Downloads\chromedriver_win32\chromedriver.exe"
    driver.get('https://stackoverflow.com/users/signup?ssrc=head&returnurl=%2fusers%2fstory%2fcurrent%27')  # signing in to google through stack overflow
    sleep(2)
    driver.find_element_by_xpath('//*[@id="openid-buttons"]/button[1]').click()  # signing in with google
    driver.find_element_by_xpath('//input[@type="email"]').send_keys(gmail_id)  # entering the gmail id
    driver.find_element_by_xpath('//*[@id="identifierNext"]').click()
    sleep(2)
    driver.find_element_by_xpath('//input[@type="password"]').send_keys(gmail_password)  # entering the password
    driver.find_element_by_xpath('//*[@id="passwordNext"]').click()
    sleep(5)
    
    driver.get('https://meet.google.com/')
    
    
    # Enter the meeting
    # Case when logged in with personal gmail account
    driver.find_element_by_css_selector('input#i3').send_keys(meet_link)  # Enter a code or link
    sleep(1)
    driver.find_element_by_css_selector('button.VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-dgl2Hf.ksBjEc.lKxP2d.cjtUbb').click()  # join
    sleep(2)
    cam_mic_selectors = driver.find_elements_by_css_selector('div.U26fgb.JRY2Pb.mUbCce.kpROve')  # camera and mic
    for e in cam_mic_selectors:
        e.click()

    sleep(2)
    driver.find_element_by_css_selector('div.uArJ5e.UQuaGc.Y5sE8d.uyXBBb.xKiqt').click()  # join now
    sleep(16)
    driver.find_element_by_css_selector('div.uArJ5e.UQuaGc.kCyAyd.QU4Gid.foXzLb.IeuGXd').click()  # participant list
    sleep(1)
    names = driver.find_elements(By.XPATH, f'.//div[contains(@class, "KV1GEc")]')  # participants
    for e in names[1:]:
        print(e.text)
        name = e.text
        name_list.append(name)
    n = int(driver.find_element_by_css_selector('div.eUyZxf span.rua5Nb').text.strip('(').strip(')'))  # no. of participants present
    no = n
    
    wb = load_workbook('Google_Attendance.xlsx')
    sheet = wb['Attendance Sheet']

    count = 0 
    for i in range(2,no+1):
        for name in name_list:
            cell = "A" + str(i)
            name_index = name_list[0]
            sheet[cell] = name_index
            count += 1
            if count >= 1:
                name_list.remove(name_index)
                break
            elif count > no:
                break 

    wb.save('Google_Attendance.xlsx')
    print(f'No. of participants : {n-1}')  # 1 participant adds up while taking attendance
    print('Attendance taken successfully!')
   
