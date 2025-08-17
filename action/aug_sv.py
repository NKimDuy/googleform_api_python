from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service

chrome_options = Options()
chrome_options.add_argument("--headless=new")  # Chế độ headless mới
chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # Tắt automation
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option("useAutomationExtension", False)
import time
import random





# Khởi tạo driver
#driver = webdriver.Chrome(options=chrome_options)
# Che giấu navigator.webdriver
# driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
# "source": """
#       Object.defineProperty(navigator, 'webdriver', {
#             get: () => undefined
#       });
# """
# })

driver = webdriver.Chrome()
driver.get('https://docs.google.com/forms/d/e/1FAIpQLSdB_NQYerrruWnjaklubqwJUt4ruKFA-GwYtC-VMd0-0cR5jQ/viewform')


try:
    try:
        text_input = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, ".//input[@type='text']"))
        )
        text_input.send_keys('111')
        time.sleep(0.5)
    except Exception as e: 
        print('không tìm thấy input')
    try:
        next_after_mssv = driver.find_element(By.XPATH, "//span[text()='Tiếp']")
        driver.execute_script("arguments[0].scrollIntoView(true);", next_after_mssv)
        next_after_mssv.click()
        time.sleep(5)
    except Exception as e:
        print('không nhấn vào được nút next')
    try:
        agree = ['Đồng ý', 'Hoàn toàn đồng ý']
        satisfy = ['Khá hài lòng', 'Khá hài lòng']

        question_list = {
            'Giảng viên (GV) giớ thiệu đề cương chi tiết và chuẩn đầu ra (CĐR) của môn học đầy đủ, rõ ràng trước khi bắt đầu môn học:': agree,
            'GV giải thích phương pháp kiểm tra, đánh giá rõ ràng (thời điểm, nội dung, phương pháp kiểm tra, đánh giá) nhằm giúp sinh viên (SV) đạt được chuẩn đầu ra:': agree,
            'GV giới thiệu nguồn tài liệu tham khảo:': agree, 
            'Tài liệu được phát kịp thời cho môn học:': agree,
            'Phương pháp truyền đạt rõ ràng, dễ hiểu nhằm giúp SV đạt được chuẩn đầu ra:': agree,
            'Cách thức giảng dạy tạo hứng thú học tập cho người học:': agree,
            'Tạo điều kiện để SV tham gia tích cực vào các hoạt động trong tiết học:': agree,
            'Nêu vấn đề để SV tham gia tích cực vào các hoạt động trong tiết học:': agree,
            'Hướng dẫn sinh viên cách tự học, tự nghiên cứu ngoài giờ học:': agree,
            'Sử dụng hiệu quả các phương tiện dạy học (máy chiếu, internet...):': agree,
            'GV quan tâm đến việc tiếp thu bài giảng của sinh viên:': agree,
            'Nội dung bài giảng được trình bày đầy đủ theo đề cương chi tiết môn học:': agree,
            'Bổ xung, cập nhật những vấn đề mới bên ngoài nội dung của giáo trình:': agree,
            'Nội dung môn học được cập nhật phù hợp với thực tiễn:': agree,
            'Thực hiện nghiêm túc giờ giấc giảng dạy, sử dụng hiệu quả thời gian lên lớp:': agree,
            'Nhiệt tình và có trách nhiệm trong giảng dạy:': agree,
            'Thể hiện tính chuẩn mực tác phong nhà giáo: trang phục, lời nới, cử chỉ:': agree,
            'Có thái độ tôn trọng người học:': agree,
            'GV có sử dụng hiệu quả công nghệ hỗ trợ giảng dạy và học tập (Hệ thống quản lý học tập LMS):': agree,
            'GV theo đúng thời khóa biểu nhà trường đã đề ra:': agree,
            'GV giảng dạy theo đúng tài liệu nhà trường đã cung cấp:': agree,
            'Thời lượng hướng dẫn/giảng dạy của môn học là phù hợp:': agree,
            'Kết quả kiểm tra giữa kỳ được GV công bố trước khi kết thúc môn học:': agree,
            'GV sử dụng nhiều hình thức kiểm tra, đánh giá để tăng độ chính xác, tin cậy, tính giá trị trong đánh giá và đáp ứng CĐR:':agree,
            'GV đánh giá công bằng và phản ánh đúng năng lực của SV theo chuẩn đầu ra (CĐR):': agree,
            'Nội dung kiểm tra phù hợp với nội dung giảng dạy và CĐR:': agree,
            'Tài liệu học tập được cung cấp đúng với thông tin ghi trên đề cương môn học:': agree,
            'Anh/chị cho biết mức độ hài lòng về chất lượng giảng dạy của giảng viên:': satisfy,
            'Anh/chị cho biết mức độ hài lòng về hiệu quả giảng dạy của giảng viên:': satisfy,
            'Nhìn chung (tổng thể), Anh/Chị cho biết mức độ hài lòng về chất lượng & hiệu quả giảng dạy của giảng viên:': satisfy
        }

        while(True):
            subject = driver.find_element(By.XPATH, f"//span[contains(., 'thông tin môn học')]")
            field_container_subject = subject.find_element(By.XPATH, "./ancestor::div[@role='listitem']")
            radio_subject = WebDriverWait(field_container_subject, 5).until(
                EC.element_to_be_clickable((By.XPATH, ".//div[@role='radio']"))
            )
            radio_subject.click()
            time.sleep(1)

            for q_key, q_answer in question_list.items():
                try:
                    question = driver.find_element(By.XPATH, f"//span[contains(., '{q_key}')]")
                    field_container_question = question.find_element(By.XPATH, "./ancestor::div[@role='listitem']")
                    radio_answers = field_container_question.find_elements(By.XPATH, ".//div[@role='radio']")
                    for answer in radio_answers:
                        answer_label = answer.get_attribute("data-value")
                        if q_answer[random.randint(0, 1)] in answer_label:
                            driver.execute_script("arguments[0].scrollIntoView(true);", answer)
                            answer.click()
                            time.sleep(1)
                            break
                    
                except:
                        print(f"không tìm thấy: {q_key}")
            try:
                check_next_or_submit = driver.find_elements(By.XPATH, "//span[text()='Tiếp']")
                if check_next_or_submit:
                    print(check_next_or_submit[0].text)
                    driver.execute_script("arguments[0].scrollIntoView(true);", check_next_or_submit[0])
                    check_next_or_submit[0].click()
                    time.sleep(5)
                else:
                    submit_button = driver.find_element(By.XPATH, "//span[text()='Gửi']")
                    print(submit_button.text)
                    driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)
                    submit_button.click()
                    time.sleep(5)  # Đợi gửi form
                    break
            except Exception as e:
                print(e)
            
    except: 
        print("không tìm thấy thông tin câu hỏi")      

except:
      print("Không truy cập được vào form")