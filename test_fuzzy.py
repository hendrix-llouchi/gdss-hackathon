import difflib

def is_duplicate(new_item, existing_item):
    b_new = new_item.get("barcode", "").strip().lower()
    b_ext = existing_item.get("barcode", "").strip().lower()
    
    if b_new and b_ext and b_new != "empty" and b_ext != "empty":
        ratio = difflib.SequenceMatcher(None, b_new, b_ext).ratio()
        if ratio > 0.85:
            return True
        if ratio < 0.85:
            return False

    w_new = new_item.get("weight", "").strip().lower()
    w_ext = existing_item.get("weight", "").strip().lower()
    if w_new and w_ext and w_new != "empty" and w_ext != "empty":
        if w_new != w_ext:
            return False 
            
    br_new = new_item.get("brand", "").strip().lower()
    br_ext = existing_item.get("brand", "").strip().lower()
    
    brand_match = False
    if not br_new or not br_ext or br_new == "empty" or br_ext == "empty":
        brand_match = True
    else:
        if difflib.SequenceMatcher(None, br_new, br_ext).ratio() > 0.8:
            brand_match = True
            
    if not brand_match:
        return False
        
    in_new = new_item.get("item_name", "").strip().lower()
    in_ext = existing_item.get("item_name", "").strip().lower()
    
    if difflib.SequenceMatcher(None, in_new, in_ext).ratio() > 0.8:
        return True
        
    return False

p1 = {"item_name": "KILIF KONCENTRAT DETERGENT - POW", "barcode": "599886628377", "brand": "KILIF"}
p2 = {"item_name": "KILIF CONCENTRATED DETERGENT PO", "barcode": "999866628377", "brand": "KIILIF"}
print("P1 vs P2:", is_duplicate(p1, p2))

p3 = {"item_name": "KIVO CLASSIC TOMATO MIX 60G", "barcode": "EMPTY", "brand": "KIVO", "weight": "60G"}
p4 = {"item_name": "KIVO CLASSIC TOMATO MIX", "barcode": "EMPTY", "brand": "KIVO", "weight": ""}
print("P3 vs P4:", is_duplicate(p3, p4))
