import tkinter as tk
import subprocess
import openpyxl
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def run_script():
    start_num = start_entry.get()
    end_num = end_entry.get()

    service = Service(executable_path="chromedriver.exe")
    driver = webdriver.Chrome(service=service)

    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    current_row = 1

    for value in range(int(start_num), int(end_num)+1):  
        driver.get("http://www.aduanet.gob.pe/cl-ad-itconsmanifiesto/manifiestoITS01Alias?accion=cargarFrmConsultaManifiestoExportacion")

        input_element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.NAME, "CMc1_Numero")))
        input_element.clear()
        input_element.send_keys(str(value))

        consultar_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Consultar']")))
        consultar_button.click()

        time.sleep(5)

        tables = driver.find_elements(By.XPATH, "//table[@width='99%'] | //table[@width='100%']")

        data = []

        for table in tables:
            rows = table.find_elements(By.TAG_NAME, "tr")

            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                row_data = [cell.text.strip() for cell in cells]
                data.append(row_data)

        for row_data in data:
            for col_idx, cell_value in enumerate(row_data, start=1):
                worksheet.cell(row=current_row, column=col_idx, value=cell_value)
            current_row += 1  

        time.sleep(2)  

    driver.quit()

    workbook.save("temp_table.xlsx")

    subprocess.Popen(["start", "temp_table.xlsx"], shell=True)

root = tk.Tk()
root.title("Input Numbers")
root.geometry("1000x300")  # Set the size of the window
root.configure(bg="white")  # Set the background color to white

# Define a larger font size
font_size = ("Arial", 16)

start_label = tk.Label(root, text="Start Number:", font=font_size, bg="white")
start_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
start_entry = tk.Entry(root, font=font_size)
start_entry.grid(row=0, column=1, padx=5, pady=5)

end_label = tk.Label(root, text="End Number:", font=font_size, bg="white")
end_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
end_entry = tk.Entry(root, font=font_size)
end_entry.grid(row=1, column=1, padx=5, pady=5)

run_button = tk.Button(root, text="Search", font=font_size, command=run_script)
run_button.grid(row=2, columnspan=2, padx=5, pady=10)

# Add the resized image to the corner
image = tk.PhotoImage(file="Pricomreit.png")
resized_image = image.subsample(2)  # Resize the image by a factor of 2
image_label = tk.Label(root, image=resized_image, bg="white")
image_label.grid(row=0, column=2, rowspan=3, padx=5, pady=5)

root.mainloop()