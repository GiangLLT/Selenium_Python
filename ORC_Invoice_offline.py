import requests

# API key miễn phí của bạn từ OCR Space
api_key = 'TEST'

# Đường dẫn đến file hình ảnh cần xử lý
image_path = '.\Data\Invoice\hd.png'

# URL của API endpoint OCR Space
api_url = 'https://api.ocr.space/parse/image'

# Mở file hình ảnh ở chế độ nhị phân
with open(image_path, 'rb') as image_file:
    # Chuẩn bị payload và files cho yêu cầu POST
    payload = {
        'apikey': api_key,
        'language': 'vie',  # 'vie' là mã ngôn ngữ cho tiếng Việt
    }
    files = {'file': image_file}

    # Thực hiện yêu cầu POST đến OCR Space API
    response = requests.post(api_url, files=files, data=payload)

    # Kiểm tra nếu yêu cầu thành công
    if response.status_code == 200:
        result = response.json()
        text = result.get("ParsedResults")[0].get("ParsedText")
        print("Nội dung nhận diện từ hóa đơn:")
        print(text)
    else:
        print("Lỗi:", response.status_code, response.text)
