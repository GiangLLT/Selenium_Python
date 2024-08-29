import requests
import json

# API key của bạn (thay thế bằng API key thực tế)
api_key = 'TEST'

# Đường dẫn đến file hình ảnh cần xử lý
image_path = '.\Data\Invoice\hd.png'

# URL của API endpoint Asprise OCR
api_url = 'https://ocr.asprise.com/api/v1/receipt'

# Chuẩn bị payload với các tham số cần thiết
payload = {
    'api_key': api_key,
    'recognizer': 'auto',  # Sử dụng 'auto' để API tự xác định bộ nhận diện tốt nhất
    'ref_no': 'ocr_python_123',  # Số tham chiếu để theo dõi
}

# Mở file hình ảnh ở chế độ nhị phân
with open(image_path, 'rb') as image_file:
    files = {'file': image_file}

    # Thực hiện yêu cầu POST đến Asprise OCR API
    response = requests.post(api_url, data=payload, files=files)

    # Kiểm tra nếu yêu cầu thành công
    if response.status_code == 200:
        # Phân tích kết quả JSON
        result = response.json()
        data = result["receipts"]
        print(json.dumps(data, indent=4, ensure_ascii=False))

        # In ra kết quả toàn bộ JSON (nếu cần kiểm tra)
        # print(json.dumps(result, indent=4, ensure_ascii=False))

        # Trích xuất và in các mục cụ thể từ kết quả
        print("Thông tin hóa đơn:")
        print(f"Ngày: {result.get('date', 'N/A')}")
        print(f"Đơn vị bán hàng: {result.get('merchant_name', 'N/A')}")
        print(f"Mã số thuế: {result.get('merchant_tax_id', 'N/A')}")
        print(f"Địa chỉ: {result.get('merchant_address', 'N/A')}")
        print(f"Số tài khoản: {result.get('merchant_account', 'N/A')}")

        print("\nThông tin người mua hàng:")
        print(f"Họ tên người mua hàng: {result.get('customer_name', 'N/A')}")
        print(f"Tên đơn vị: {result.get('customer_company', 'N/A')}")
        print(f"Mã số thuế: {result.get('customer_tax_id', 'N/A')}")
        print(f"Địa chỉ: {result.get('customer_address', 'N/A')}")

        print("\nChi tiết hàng hóa:")
        for item in result.get('items', []):
            print(f"STT: {item.get('line_no', 'N/A')}")
            print(f"Mã hàng: {item.get('sku', 'N/A')}")
            print(f"Tên hàng hóa: {item.get('description', 'N/A')}")
            print(f"Đơn vị tính: {item.get('unit', 'N/A')}")
            print(f"Số lượng: {item.get('quantity', 'N/A')}")
            print(f"Đơn giá: {item.get('unit_price', 'N/A')}")
            print(f"Thành tiền: {item.get('total', 'N/A')}")
            print("-" * 30)

        print("\nThông tin thanh toán:")
        print(f"Cộng tiền hàng: {result.get('subtotal', 'N/A')}")
        print(f"Thuế suất GTGT: {result.get('tax_rate', 'N/A')}")
        print(f"Tiền thuế GTGT: {result.get('tax', 'N/A')}")
        print(f"Tổng tiền thanh toán: {result.get('total', 'N/A')}")
        print(f"Số tiền viết bằng chữ: {result.get('total_in_words', 'N/A')}")
    else:
        print("Lỗi:", response.status_code, response.text)
