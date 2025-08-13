import pandas as pd
import re
import os
import openpyxl # Đảm bảo thư viện này đã được cài đặt: pip install openpyxl

def convert_amazon_feed():
    # --- Định nghĩa các header và tag để đảm bảo thứ tự và tên cột chính xác ---
    # US Human Headers (Thứ tự cột cuối cùng mong muốn cho hàng hiển thị của file US - Hàng 2 Output)
    us_human_headers_output_order = [
        "Product Type", "Seller SKU", "Brand Name", "Product Name",
        "Item Type Keyword", "Department",
        "Shirt Body Type", "Shirt Height Type", "Fit Type", "NeckStyle", "Style", "List Price", "Material type",
        "Standard Price", "Quantity",
        "Main Image URL", "Other Image Url1", "Other Image Url2", "Other Image Url3",
        "Other Image Url4", "Other Image Url5", "Other Image Url6", "Other Image Url7",
        "Other Image Url8", "Product Description", "Key Product Features",
        "Key Product Features", "Key Product Features", "Key Product Features",
        "Key Product Features", "Search Terms", "Shipping-Template", "Color",
        "Color Map", "Size", "Size Map", "Handling Time", "Target Gender",
        "Age Range Description", "Shirt Size System", "Shirt Size Class",
        "Shirt Size Value", "Adult Flag", "Fabric Type", "Sleeve Type",
        "Outer Material Type"
    ]

    # US Tags (Các tên cột độc nhất cho hàng "tag" của file US - Hàng 3 Output)
    us_tags_internal_df_cols = [
        "feed_product_type", "item_sku", "brand_name", "item_name",
        "item_type_keyword", "department_name",
        "shirt_body_type", "shirt_height_type", "fit_type", "neck_style", "style_name", "list_price", "material_type",
        "standard_price", "quantity",
        "main_image_url", "other_image_url1", "other_image_url2", "other_image_url3",
        "other_image_url4", "other_image_url5", "other_image_url6", "other_image_url7",
        "other_image_url8", "product_description", "bullet_point1",
        "bullet_point2", "bullet_point3", "bullet_point4",
        "bullet_point5", "generic_keywords", "merchant_shipping_group_name", "color_name",
        "color_map", "size_name", "size_map", "fulfillment_latency", "target_gender",
        "age_range_description", "shirt_size_system", "shirt_size_class",
        "shirt_size", "is_adult_product", "fabric_type", "sleeve_type",
        "outer_material_type"
    ]

    # Hardcoded instruction rows for Amazon US template
    amazon_us_instructions_row1 = [
        "TemplateType=fptcustom",
        "Version=2021.0321",
        "TemplateSignature=U0hJUlQ=",
        "The top three rows are for Amazon.com use only. Do not modify or delete the top three rows."
    ]

    # --- Blacklist Keywords (DANH SÁCH BLACKLIST MỚI) ---
    blacklist_keywords = [
        "Whyitsme", "Cottagecore", "Trump", "Biden", "Reggae", "Smoke Daddy", "Celtic Cross", "Bob Marley", "Family Guy", "Gay Cat", "Gay Trash", "Jockey", "Fishy", "Venom", "Boba", "BSN", "Uterus", "Van Gogh",
        "CARHARTT", "Nonni", "Kangaroo", "Tuxedo", "Dibble", "Dabble", "Oh ship", "comica", "COHIBA", "Jurassic", "Jeep", "Jeeps", "Rubiks", "Adventure Before Dementia", "antisocial", "anti social", "Cobra", "Python",
        "Spirit Halloween", "Got Titties", "Le Tits Now", "Mack Trucks", "V-buck", "V buck", "Vbuck", "World Traveler", "Rollerblade", "Black Lives Matter", "Just The Tip", "In My Defense", "Van Gogh",
        "U.S.Army", "US Army", "Crazy Chicken Lady", "Christmas In July", "Grill Sergeant", "Ducks Unlimited", "SOTALLY Tober", "Birds aren't Real", "Pickleballer", "Quaker", "Vampire Mansion",
        "Lampoon's", "Lampoons", "Lampoon", "krampus", "griswold", "Brainrot", "Disney", "Marvel", "Star Wars", "Music Television", "MTV", "Fender", "Nightmare Before Christmas", "Life is Good",
        "WWE", "NFL", "NBA", "Robux", "ASPCA",
    ]
    # Chuyển tất cả từ khóa về chữ thường để so sánh không phân biệt hoa thường
    blacklist_keywords_lower = [k.lower() for k in blacklist_keywords]

    # --- Keyword Replacements (DANH SÁCH THAY THẾ TỪ KHÓA MỚI) ---
    keyword_replacements = {
        "Guess": "Funny",
        "Sakura": "Flower",
        "Superhero": "Heroes",
        "Yeti": "Bigfoot",
        "Beast": "Strong",
        "Diesel": "Handyman",
        "K-Pop": "Korean Music",
        "Kpop": "Korean Music",
        "Frisbee": "Sport",
        "Coach": "Fun",
        "KOOZIE": "Drinking",
        "Prosecco": "Drinking",
        "Craftsman": "Handyman",
        "Pajama": "Costume",
        "Pajamas": "Costume",
        "Shark Week": "Shark Lovers",
        "BANNED": "Reading Lover",
        "Arcade": "Game Machine",
        "Ducky": "Duck Lovers",
        "Skittles": "Fruit Candy",
    }

    # --- User Inputs ---
    input_file_name = input("Nhập tên file Excel Canada (không cần '.xlsx'): ")
    # Tự động thêm .xlsx nếu người dùng không nhập
    if not input_file_name.endswith(".xlsx"):
        input_file_path = input_file_name + ".xlsx"
    else:
        input_file_path = input_file_name

    new_prefix = input("Nhập prefix mới để thay thế (ví dụ: XYZ): ")
    while True:
        try:
            new_price = float(input("Nhập giá sản phẩm mới (ví dụ: 19.99): "))
            break
        except ValueError:
            print("Giá không hợp lệ. Vui lòng nhập một số.")

    print(f"\nĐang xử lý file: {input_file_path}")

    if not os.path.exists(input_file_path):
        print(f"LỖI: Không tìm thấy file '{input_file_path}'. Vui lòng kiểm tra lại tên và đường dẫn file.")
        print(f"Đảm bảo file '{input_file_path}' nằm cùng thư mục với script Python này hoặc cung cấp đường dẫn đầy đủ.")
        return

    try:
        # Đọc toàn bộ sheet mà không chỉ định header hay skiprows
        df_raw = pd.read_excel(input_file_path, header=None, engine='openpyxl')

        # Xác định hàng tags (Hàng 3 trong Excel, index 2 trong DataFrame 0-indexed)
        original_canada_tags_from_file = [str(x) for x in df_raw.iloc[2].dropna().tolist() if str(x).strip() != '']

        # Lấy dữ liệu từ hàng 4 trở đi (index 3 trong DataFrame)
        df_data = df_raw.iloc[3:].copy()
        df_data.columns = original_canada_tags_from_file[:len(df_data.columns)] # Gán tags làm tên cột. Cắt tags nếu số cột data ít hơn tags.

        # Kiểm tra và xử lý trường hợp số cột dữ liệu nhiều hơn số tags
        if len(df_data.columns) < len(original_canada_tags_from_file):
            missing_cols_in_data = original_canada_tags_from_file[len(df_data.columns):]
            for col in missing_cols_in_data:
                df_data[col] = pd.NA

        if df_data.empty:
            print("CẢNH BÁO: Không có dữ liệu trong file nhập sau khi xác định header. Tạo DataFrame rỗng để tiếp tục.")
            df_data = pd.DataFrame(columns=original_canada_tags_from_file)

    except Exception as e:
        print(f"LỖI: Khi đọc file Excel hoặc xác định cấu trúc: {e}")
        print("Vui lòng kiểm tra lại file đầu vào của bạn có đúng cấu trúc (Hàng 1: instructions, Hàng 2: Human Headers, Hàng 3: Internal Tags) không.")
        return

    # --- Perform data transformations ---
    if not df_data.empty: # Chỉ thực hiện biến đổi nếu có dữ liệu
        # 1. Rename columns (internal tags)
        internal_tag_rename_mapping = {
            "recommended_browse_nodes": "item_type_keyword",
            "material_composition": "fabric_type"
        }
        actual_rename_mapping = {k: v for k, v in internal_tag_rename_mapping.items() if k in df_data.columns}
        if actual_rename_mapping:
            df_data = df_data.rename(columns=actual_rename_mapping)

        # 2. Update column values
        if "item_type_keyword" in df_data.columns:
            df_data["item_type_keyword"] = df_data["item_type_keyword"].astype(str).replace("17275638011", "fashion-t-shirts")

        if "shirt_size_system" in df_data.columns:
            df_data["shirt_size_system"] = df_data["shirt_size_system"].astype(str).replace("CA/US", "US")

        # 3. Add new column 'sleeve_type'
        if "sleeve_type" not in df_data.columns:
            df_data["sleeve_type"] = "Short Sleeve"
        else:
            df_data["sleeve_type"] = "Short Sleeve"

        # 4. Replace Prefix in Seller SKU and Product Name
        old_prefix = None
        if "item_sku" in df_data.columns and not df_data["item_sku"].empty:
            first_valid_sku_idx = df_data["item_sku"].first_valid_index()
            if first_valid_sku_idx is not None and pd.notna(df_data.loc[first_valid_sku_idx, "item_sku"]):
                first_sku = str(df_data.loc[first_valid_sku_idx, "item_sku"])
                match = re.match(r"^([A-Za-z0-9]+)\d{13,}-\w+", first_sku)
                if match:
                    old_prefix = match.group(1)
            
        if old_prefix and old_prefix != new_prefix:
            if "item_sku" in df_data.columns:
                df_data["item_sku"] = df_data["item_sku"].astype(str).apply(
                    lambda x: re.sub(r"^" + re.escape(old_prefix), new_prefix, x) if re.match(r"^" + re.escape(old_prefix), x) else x
                )
            if "item_name" in df_data.columns:
                df_data["item_name"] = df_data["item_name"].astype(str).apply(
                    lambda x: x.replace(old_prefix, new_prefix) if old_prefix in x else x
                )
        elif old_prefix == new_prefix:
            pass # No need to update if prefix is same

        # 5. Price Update
        if "standard_price" in df_data.columns:
            df_data["standard_price"] = new_price
        
        if "list_price" in df_data.columns:
            df_data["list_price"] = new_price
        
        # --- NEW: Replace Blacklist Keywords ---
        print("\nĐang thay thế các từ khóa theo danh sách...")
        df_data = replace_keywords(df_data, keyword_replacements)
        print("Đã hoàn thành việc thay thế từ khóa.")

        # --- NEW: Filter Blacklist Rows ---
        print("\nĐang lọc các hàng chứa từ khóa trong blacklist...")
        initial_rows = len(df_data)
        df_data = filter_blacklist_rows(df_data, blacklist_keywords_lower)
        rows_deleted = initial_rows - len(df_data)
        if rows_deleted > 0:
            print(f"Đã xóa {rows_deleted} hàng do chứa từ khóa trong blacklist.")
        else:
            print("Không có hàng nào bị xóa do không chứa từ khóa blacklist.")

    else:
        print("CẢNH BÁO: DataFrame dữ liệu rỗng. Không thực hiện biến đổi nào.")

    # --- Column Reordering for output ---
    # Ensure all target US columns exist in df_data, add as NA if missing
    for col in us_tags_internal_df_cols:
        if col not in df_data.columns:
            df_data[col] = pd.NA

    df_converted = df_data[us_tags_internal_df_cols]

    # --- Save the converted file using openpyxl for precise header control ---
    base_name = os.path.splitext(input_file_path)[0]
    output_file_path = f"_Output_{base_name}.xlsx"

    try:
        # Create a brand new workbook and sheet
        workbook = openpyxl.Workbook()
        sheet = workbook.active

        # Determine max columns for padding. Use the longest expected header/tag list
        MIN_COLS = max(
            len(us_human_headers_output_order),
            len(us_tags_internal_df_cols)
        ) + 5 # Add a buffer of 5 columns for safety

        # Helper function to pad lists with empty strings and convert all elements to string
        def pad_and_str_list(lst, target_len):
            padded = [str(x) if pd.notna(x) else '' for x in lst]
            return padded + [''] * (target_len - len(padded))

        # 1. Ghi hàng hướng dẫn đầu tiên (Hàng 1 Excel)
        padded_row_1 = pad_and_str_list(amazon_us_instructions_row1, MIN_COLS)
        sheet.append(padded_row_1)

        # 2. Ghi Human Headers (Hàng 2 Excel)
        padded_human_headers = pad_and_str_list(us_human_headers_output_order, MIN_COLS)
        sheet.append(padded_human_headers)

        # 3. Ghi US Tags (Hàng 3 Excel)
        padded_us_tags = pad_and_str_list(us_tags_internal_df_cols, MIN_COLS)
        sheet.append(padded_us_tags)

        # 4. Append the DataFrame data (Từ Hàng 4 Excel trở đi)
        if not df_converted.empty:
            for row_data in df_converted.values.tolist():
                padded_data_row = pad_and_str_list(row_data, MIN_COLS)
                sheet.append(padded_data_row)
        else:
            print("CẢNH BÁO: Không có dữ liệu sản phẩm nào để ghi vào file Excel.")

        workbook.save(output_file_path)

        print(f"\nChuyển đổi thành công! File mới đã được lưu tại: {output_file_path}")
        print("Vui lòng kiểm tra lại file đã chuyển đổi trước khi upload lên Amazon US.")
    except Exception as e:
        print(f"\nLỖI KHI LƯU FILE EXCEL: {e}")
        print("Vui lòng đảm bảo file Excel đầu ra không đang mở trong các ứng dụng khác và thử lại.")

# --- NEW FUNCTION: Replace Keywords ---
def replace_keywords(df, replacements):
    """
    Thay thế các từ khóa trong DataFrame bằng các từ khóa mới theo từ điển.
    Thực hiện thay thế không phân biệt hoa thường.
    """
    if df.empty:
        return df
    
    df_copy = df.copy()
    
    # Duyệt qua từng cặp từ khóa cũ và từ khóa mới
    for old_keyword, new_keyword in replacements.items():
        # Tạo regex pattern để tìm kiếm từ khóa cũ một cách không phân biệt hoa thường
        # Sử dụng re.escape để xử lý các ký tự đặc biệt
        pattern = re.compile(re.escape(old_keyword), re.IGNORECASE)
        
        # Áp dụng thay thế trên tất cả các cột kiểu chuỗi
        for col in df_copy.columns:
            if pd.api.types.is_string_dtype(df_copy[col]):
                df_copy[col] = df_copy[col].astype(str).apply(
                    lambda x: pattern.sub(new_keyword, x)
                )
    return df_copy

# --- NEW FUNCTION: Filter Blacklist Rows ---
def filter_blacklist_rows(df, blacklist_keywords_lower):
    """
    Lọc và xóa các hàng trong DataFrame nếu bất kỳ ô nào trong hàng đó chứa từ khóa trong blacklist.
    So sánh không phân biệt hoa thường.
    """
    if df.empty:
        return df

    # Tạo một Series boolean, mặc định tất cả là True (không bị blacklist)
    # Nếu bất kỳ ô nào trong hàng chứa từ khóa blacklist, hàng đó sẽ bị đánh dấu là False
    rows_to_keep = pd.Series(True, index=df.index)

    for keyword in blacklist_keywords_lower:
        # Kiểm tra từng cột trong DataFrame
        for col in df.columns:
            # Chuyển đổi tất cả giá trị trong cột về chuỗi để tìm kiếm, xử lý NaN an toàn
            # Sử dụng str.contains với regex=False để tìm kiếm chuỗi con đơn giản
            # và case=False để không phân biệt hoa thường
            if pd.api.types.is_string_dtype(df[col]):
                # Chỉ kiểm tra nếu cột là kiểu chuỗi để tránh lỗi
                mask = df[col].astype(str).str.lower().str.contains(re.escape(keyword), na=False, regex=False)
            else:
                # Đối với các kiểu dữ liệu khác, chuyển sang chuỗi và kiểm tra
                mask = df[col].astype(str).str.lower().str.contains(re.escape(keyword), na=False, regex=False)
            
            # Nếu có bất kỳ hàng nào phù hợp với từ khóa này, đánh dấu rows_to_keep của hàng đó là False
            rows_to_keep = rows_to_keep & (~mask) # Giữ lại những hàng KHÔNG chứa từ khóa

    # Trả về DataFrame chỉ với các hàng được giữ lại
    return df[rows_to_keep]


# Call the main function to run the script
if __name__ == "__main__":
    convert_amazon_feed()