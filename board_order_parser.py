import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import inch, landscape
from reportlab.lib import colors
from pypdf import PdfReader, PdfWriter
import zipfile

# --------------------------------------
# Page Configuration
# --------------------------------------
st.set_page_config(
    page_title="Board Order Manager",
    page_icon="🪵",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --------------------------------------
# Dark Mode Custom CSS Styling
# --------------------------------------
st.markdown("""
<style>
    /* Dark Mode Base */
    .main {
        background: #0f1419;
        color: #e4e6eb;
    }
    
    .stApp {
        background: #0f1419;
    }
    
    /* Sidebar Dark Styling */
    [data-testid="stSidebar"] {
        background: #1a1f2e;
        border-right: 1px solid #2d3748;
    }
    
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
        color: #e4e6eb;
    }
    
    /* Sidebar Navigation Links */
    .nav-link {
        display: block;
        padding: 12px 15px;
        margin: 4px 0;
        border-radius: 10px;
        color: #a0aec0;
        text-decoration: none;
        transition: all 0.2s ease;
        cursor: pointer;
    }
    
    .nav-link:hover {
        background: #2d3748;
        color: #e4e6eb;
        text-decoration: none;
    }
    
    .nav-link.active {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
    }
    
    /* Metric Cards Dark */
    [data-testid="stMetric"] {
        background: linear-gradient(135deg, #1e2432 0%, #252d3d 100%);
        border: 1px solid #2d3748;
        padding: 25px 20px;
        border-radius: 16px;
        border-left: 3px solid #667eea;
    }
    
    [data-testid="stMetric"]:hover {
        transform: translateY(-3px);
        box-shadow: 0 10px 30px rgba(102, 126, 234, 0.2);
        transition: all 0.3s ease;
        border-color: #667eea;
    }
    
    [data-testid="stMetric"] label {
        font-size: 0.85em !important;
        color: #a0aec0 !important;
        font-weight: 600 !important;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    [data-testid="stMetric"] [data-testid="stMetricValue"] {
        font-size: 2.5em !important;
        font-weight: 700 !important;
        color: #e4e6eb !important;
    }
    
    /* Headers Dark */
    h1 {
        color: #e4e6eb;
        font-weight: 700;
        padding-bottom: 15px;
        border-bottom: 3px solid #667eea;
        margin-bottom: 30px;
    }
    
    h2 {
        color: #e4e6eb;
        font-weight: 600;
        margin-top: 40px;
        margin-bottom: 20px;
    }
    
    h3 {
        color: #cbd5e0;
        font-weight: 600;
        margin-bottom: 15px;
    }
    
    /* Buttons Dark */
    .stButton button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 12px 24px;
        font-weight: 600;
        transition: all 0.3s ease;
        width: 100%;
    }
    
    .stButton button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 20px rgba(102, 126, 234, 0.4);
    }
    
    /* File Uploader Dark */
    [data-testid="stFileUploader"] {
        background: #1a1f2e;
        padding: 40px;
        border-radius: 12px;
        border: 2px dashed #2d3748;
    }
    
    [data-testid="stFileUploader"]:hover {
        border-color: #667eea;
        background: #1e2432;
    }
    
    [data-testid="stFileUploader"] label {
        color: #e4e6eb !important;
    }
    
    [data-testid="stFileUploader"] section {
        border-color: #2d3748 !important;
    }
    
    /* Info boxes Dark */
    .stAlert {
        background: linear-gradient(135deg, #667eea20, #764ba220) !important;
        border: 1px solid #667eea40 !important;
        border-radius: 10px;
        border-left: 4px solid #667eea !important;
        color: #cbd5e0 !important;
    }
    
    /* Success boxes */
    [data-baseweb="notification"] {
        background: #1a1f2e !important;
        border: 1px solid #48bb78 !important;
        color: #e4e6eb !important;
    }
    
    /* Dataframe Dark */
    [data-testid="stDataFrame"] {
        border-radius: 12px;
        overflow: hidden;
    }
    
    [data-testid="stDataFrame"] table {
        background: #1a1f2e !important;
        color: #e4e6eb !important;
    }
    
    [data-testid="stDataFrame"] thead tr th {
        background: #2d3748 !important;
        color: #e4e6eb !important;
    }
    
    [data-testid="stDataFrame"] tbody tr {
        background: #1e2432 !important;
        color: #cbd5e0 !important;
    }
    
    [data-testid="stDataFrame"] tbody tr:hover {
        background: #252d3d !important;
    }
    
    /* Expander Dark */
    [data-testid="stExpander"] {
        background: #1a1f2e !important;
        border-radius: 10px;
        border: 1px solid #2d3748 !important;
        margin-bottom: 10px;
    }
    
    [data-testid="stExpander"] [data-testid="stMarkdownContainer"] {
        color: #e4e6eb !important;
    }
    
    /* Progress bar */
    .stProgress > div > div {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
    }
    
    /* Download button */
    .stDownloadButton button {
        background: linear-gradient(135deg, #48bb78 0%, #38a169 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 10px 20px;
        font-weight: 600;
        width: 100%;
    }
    
    .stDownloadButton button:hover {
        transform: translateY(-1px);
        box-shadow: 0 5px 15px rgba(72, 187, 120, 0.4);
    }
    
    /* Text color overrides */
    p, span, div {
        color: #cbd5e0;
    }
    
    strong {
        color: #e4e6eb;
    }
    
    /* Section divider */
    hr {
        border: none;
        border-top: 1px solid #2d3748;
        margin: 40px 0;
    }
    
    /* Spinner Dark */
    .stSpinner > div {
        border-top-color: #667eea !important;
    }
    
    /* Input fields */
    input, textarea, select {
        background: #1a1f2e !important;
        color: #e4e6eb !important;
        border: 1px solid #2d3748 !important;
    }
    
    /* Markdown text */
    .stMarkdown {
        color: #cbd5e0 !important;
    }
    
    /* Status indicator */
    .status-indicator {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        background: #2d3748;
        padding: 6px 12px;
        border-radius: 20px;
        font-size: 0.85em;
    }
    
    .status-dot {
        width: 8px;
        height: 8px;
        background: #48bb78;
        border-radius: 50%;
        animation: pulse 2s infinite;
    }
    
    @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.5; }
    }
</style>
""", unsafe_allow_html=True)

# --------------------------------------
# Helper Functions
# --------------------------------------
def clean_text(s: str) -> str:
    """Cleans unwanted symbols and formatting."""
    if not s:
        return ""
    s = re.sub(r"■|Seller Name|Your Orders|Returning your item:", "", s)
    s = re.sub(r"\(Most popular\)", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s{2,}", " ", s)
    return s.strip()

def extract_design_number(text: str) -> str:
    """Extract design number from the customization text."""
    # Check for "No Engraving" first
    if re.search(r"No\s*Engraving.*(?:Blank\s*Board)?", text, re.IGNORECASE):
        return "NO_DESIGN"
    
    design_match = re.search(r"Design\s*#?:?\s*Design\s*(\d+)", text, re.IGNORECASE)
    if design_match:
        return design_match.group(1)
    
    design_match = re.search(r"Design\s*#?:?\s*(\d+)", text, re.IGNORECASE)
    if design_match:
        return design_match.group(1)
    
    return ""

def extract_engraving_type(text: str) -> str:
    """Extract the engraving type from the order."""
    if re.search(r"Board\s*\+\s*Utensils\s*Engraving", text, re.IGNORECASE):
        return "Board+Utensils Engraving"
    elif re.search(r"Only\s*Board\s*Engraving", text, re.IGNORECASE):
        return "Only Board Engraving"
    elif re.search(r"No\s*Engraving", text, re.IGNORECASE):
        return "No Engraving"
    return ""

def generate_manufacturing_labels(dataframe):
    """Generate 4x6 manufacturing labels for board orders."""
    buf = BytesIO()
    page_size = landscape((4 * inch, 6 * inch))
    c = canvas.Canvas(buf, pagesize=page_size)
    W, H = page_size
    left = 0.3 * inch
    right = W - 0.3 * inch
    top = H - 0.3 * inch

    for _, row in dataframe.iterrows():
        y = top
        
        # Header with Order ID and Quantity
        c.setFont("Helvetica-Bold", 14)
        c.drawString(left, y, f"Order ID: {row['Order ID']}")
        c.drawRightString(right, y, f"Qty: {row['Quantity']}")
        y -= 0.25 * inch
        
        # Buyer and Date
        c.setFont("Helvetica", 14)
        c.drawString(left, y, f"Buyer: {row['Buyer Name']}")
        c.drawRightString(right, y, f"Date: {row['Order Date']}")
        y -= 0.3 * inch

        # Design Number Box
        box_height = 0.6 * inch
        box_y = y - box_height
        c.setStrokeColor(colors.black)
        c.setLineWidth(3)
        c.rect(left, box_y, right - left, box_height, stroke=1, fill=0)
        
        c.setFont("Helvetica-Bold", 20)
        text_y = box_y + box_height / 2 - 0.1 * inch
        
        # Display design number or "NO DESIGN / BLANK" for blank boards
        if row['Design Number'] == "NO_DESIGN":
            c.drawCentredString(W / 2, text_y, "NO DESIGN / BLANK BOARD")
        else:
            c.drawCentredString(W / 2, text_y, f"DESIGN #{row['Design Number']}")
        
        y = box_y - 0.3 * inch

        # Engraving Type
        c.setFont("Helvetica-Bold", 14)
        c.drawString(left, y, f"Type: {row['Engraving Type']}")
        y -= 0.3 * inch

        # Board Customization Note
        c.setFont("Helvetica-Bold", 16)
        c.drawString(left, y, "Customization:")
        y -= 0.25 * inch
        
        c.setFont("Helvetica", 13)
        customization = row['Board Customization Note']
        
        # Word wrap for long customizations
        words = customization.split()
        lines = []
        current_line = []
        max_width = right - left - 0.2 * inch
        
        for word in words:
            test_line = ' '.join(current_line + [word])
            if c.stringWidth(test_line, "Helvetica", 13) < max_width:
                current_line.append(word)
            else:
                if current_line:
                    lines.append(' '.join(current_line))
                current_line = [word]
        
        if current_line:
            lines.append(' '.join(current_line))
        
        for line in lines[:3]:  # Max 3 lines for customization
            c.drawString(left + 0.1 * inch, y, line)
            y -= 0.2 * inch
        
        y -= 0.1 * inch

        # Utensil Letter (if applicable)
        if row['Utensil Letter']:
            c.setFont("Helvetica-Bold", 14)
            c.drawString(left, y, f"Utensil Letter: {row['Utensil Letter']}")
            y -= 0.3 * inch

        # Gift Note indicator
        if row['Gift Note'] == "YES":
            c.setFont("Helvetica-Bold", 14)
            c.setFillColor(colors.red)
            c.drawString(left, y, "⚠️ GIFT MESSAGE INCLUDED")
            c.setFillColor(colors.black)

        c.showPage()

    c.save()
    buf.seek(0)
    return buf

def generate_gift_message_labels(dataframe):
    """Generate gift message labels in same format as blanket orders."""
    buf = BytesIO()
    page_size = landscape((4 * inch, 6 * inch))
    c = canvas.Canvas(buf, pagesize=page_size)
    W, H = page_size

    gift_orders = dataframe[dataframe['Gift Message'] != ""]

    if len(gift_orders) == 0:
        c.setFont("Helvetica", 14)
        c.drawCentredString(W / 2, H / 2, "No gift messages found in orders")
        c.showPage()
    else:
        for _, row in gift_orders.iterrows():
            # Border
            c.setStrokeColor(colors.black)
            c.setLineWidth(3)
            c.rect(0.4 * inch, 0.4 * inch, W - 0.8 * inch, H - 0.8 * inch, stroke=1, fill=0)

            c.setFont("Times-BoldItalic", 18)
            message = row['Gift Message']
            
            # Word wrap
            words = message.split()
            lines = []
            current_line = []
            max_width = W - 1.2 * inch
            
            for word in words:
                test_line = ' '.join(current_line + [word])
                if c.stringWidth(test_line, "Times-BoldItalic", 18) < max_width:
                    current_line.append(word)
                else:
                    if current_line:
                        lines.append(' '.join(current_line))
                    current_line = [word]
            
            if current_line:
                lines.append(' '.join(current_line))

            total_height = len(lines) * 0.3 * inch
            y = (H + total_height) / 2

            for line in lines:
                c.drawCentredString(W / 2, y, line)
                y -= 0.3 * inch

            c.showPage()

    c.save()
    buf.seek(0)
    return buf

def generate_csv_files_by_design(dataframe):
    """Generate separate CSV files for each design number and return as a zip."""
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        # Group by design number
        design_groups = dataframe.groupby('Design Number')
        
        for design_num, group_df in design_groups:
            # Create CSV for this design
            csv_buffer = BytesIO()
            
            # Select and order columns for CSV
            csv_df = group_df[[
                'Order ID',
                'Order Date',
                'Buyer Name',
                'Quantity',
                'Engraving Type',
                'Board Customization Note',
                'Manual Field 1',
                'Manual Field 2',
                'Manual Field 3',
                'Utensil Letter',
                'Gift Note',
                'Gift Message'
            ]].copy()
            
            csv_df.to_csv(csv_buffer, index=False)
            csv_buffer.seek(0)
            
            # Add to zip file
            zip_file.writestr(f'Design_{design_num}.csv', csv_buffer.getvalue())
    
    zip_buffer.seek(0)
    return zip_buffer

def merge_shipping_and_manufacturing_labels(shipping_pdf_bytes, manufacturing_pdf_bytes, order_dataframe):
    """
    Merge shipping labels with manufacturing labels.
    Handles orders with multiple items (one shipping label, multiple manufacturing labels).
    """
    try:
        shipping_pdf = PdfReader(shipping_pdf_bytes)
        manufacturing_pdf = PdfReader(manufacturing_pdf_bytes)
        
        # Preserve the order of orders as they appear in the dataframe
        seen_orders = []
        order_item_counts = []
        
        for order_id in order_dataframe['Order ID']:
            if order_id not in seen_orders:
                seen_orders.append(order_id)
                item_count = len(order_dataframe[order_dataframe['Order ID'] == order_id])
                order_item_counts.append(item_count)
        
        # Build mapping: shipping label index -> list of manufacturing label indices
        shipping_to_mfg = {}
        mfg_index = 0
        
        for shipping_index, item_count in enumerate(order_item_counts):
            shipping_to_mfg[shipping_index] = list(range(mfg_index, mfg_index + item_count))
            mfg_index += item_count
        
        # Create merged PDF
        output_pdf = PdfWriter()
        total_shipping_labels = len(shipping_to_mfg)
        
        for ship_idx in range(total_shipping_labels):
            if ship_idx >= len(shipping_pdf.pages):
                break
                
            output_pdf.add_page(shipping_pdf.pages[ship_idx])
            
            if ship_idx in shipping_to_mfg:
                for mfg_idx in shipping_to_mfg[ship_idx]:
                    if mfg_idx < len(manufacturing_pdf.pages):
                        output_pdf.add_page(manufacturing_pdf.pages[mfg_idx])
        
        output_buffer = BytesIO()
        output_pdf.write(output_buffer)
        output_buffer.seek(0)
        
        return output_buffer, len(shipping_to_mfg), sum(len(v) for v in shipping_to_mfg.values())
        
    except Exception as e:
        st.error(f"Error merging labels: {str(e)}")
        return None, 0, 0

def merge_labels_by_design(shipping_pdf_bytes, manufacturing_pdf_bytes, order_dataframe, target_design=None, mixed_only=False):
    """
    Merge shipping and manufacturing labels grouped by design number.
    
    Args:
        shipping_pdf_bytes: Shipping labels PDF
        manufacturing_pdf_bytes: Manufacturing labels PDF
        order_dataframe: DataFrame with all orders
        target_design: Specific design number to merge (e.g., "1", "2", etc.)
        mixed_only: If True, only merge orders with mixed designs
    
    Returns:
        Merged PDF buffer, number of orders, number of items
    """
    try:
        shipping_pdf = PdfReader(shipping_pdf_bytes)
        manufacturing_pdf = PdfReader(manufacturing_pdf_bytes)
        
        # Identify pure vs mixed design orders
        order_design_map = {}  # order_id -> list of design numbers
        for order_id in order_dataframe['Order ID'].unique():
            order_items = order_dataframe[order_dataframe['Order ID'] == order_id]
            designs = order_items['Design Number'].unique().tolist()
            order_design_map[order_id] = designs
        
        # Build list of orders in original order
        seen_orders = []
        for order_id in order_dataframe['Order ID']:
            if order_id not in seen_orders:
                seen_orders.append(order_id)
        
        # Filter orders based on target_design or mixed_only
        filtered_orders = []
        for order_id in seen_orders:
            designs = order_design_map[order_id]
            is_mixed = len(designs) > 1
            
            if mixed_only:
                # Only include mixed design orders
                if is_mixed:
                    filtered_orders.append(order_id)
            elif target_design:
                # Only include pure orders with matching design
                if not is_mixed and target_design in designs:
                    filtered_orders.append(order_id)
            else:
                # Include all orders
                filtered_orders.append(order_id)
        
        if not filtered_orders:
            return None, 0, 0
        
        # Create mapping of order indices
        order_to_shipping_idx = {}
        order_to_mfg_indices = {}
        
        shipping_idx = 0
        mfg_idx = 0
        
        for order_id in seen_orders:
            order_to_shipping_idx[order_id] = shipping_idx
            shipping_idx += 1
            
            # Get manufacturing label indices for this order
            item_count = len(order_dataframe[order_dataframe['Order ID'] == order_id])
            order_to_mfg_indices[order_id] = list(range(mfg_idx, mfg_idx + item_count))
            mfg_idx += item_count
        
        # Create merged PDF for filtered orders
        output_pdf = PdfWriter()
        
        for order_id in filtered_orders:
            ship_idx = order_to_shipping_idx[order_id]
            mfg_indices = order_to_mfg_indices[order_id]
            
            # Add shipping label
            if ship_idx < len(shipping_pdf.pages):
                output_pdf.add_page(shipping_pdf.pages[ship_idx])
            
            # Add manufacturing labels
            for mfg_idx in mfg_indices:
                if mfg_idx < len(manufacturing_pdf.pages):
                    output_pdf.add_page(manufacturing_pdf.pages[mfg_idx])
        
        # Write to buffer
        output_buffer = BytesIO()
        output_pdf.write(output_buffer)
        output_buffer.seek(0)
        
        # Calculate totals
        total_items = sum(len(order_to_mfg_indices[order_id]) for order_id in filtered_orders)
        
        return output_buffer, len(filtered_orders), total_items
        
    except Exception as e:
        st.error(f"Error merging labels by design: {str(e)}")
        return None, 0, 0

# --------------------------------------
# SIDEBAR WITH FUNCTIONAL NAVIGATION
# --------------------------------------
with st.sidebar:
    st.markdown("# 🪵 Board Manager")
    st.markdown("### Version 1.1 Dark")
    st.markdown("---")
    
    st.markdown("#### 📋 Quick Navigation")
    
    st.markdown('<a href="#upload-order" class="nav-link">📄 Upload Order</a>', unsafe_allow_html=True)
    st.markdown('<a href="#dashboard" class="nav-link">📊 Dashboard</a>', unsafe_allow_html=True)
    st.markdown('<a href="#design-breakdown" class="nav-link">🎨 Design Breakdown</a>', unsafe_allow_html=True)
    st.markdown('<a href="#generate-files" class="nav-link">📥 Generate Files</a>', unsafe_allow_html=True)
    st.markdown('<a href="#label-merge" class="nav-link">🔄 Label Merge</a>', unsafe_allow_html=True)
    st.markdown('<a href="#design-merge" class="nav-link">🎨 Merge by Design</a>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    st.markdown("#### ✨ Features")
    st.markdown("✓ PDF Parsing")
    st.markdown("✓ Design-Specific CSVs")
    st.markdown("✓ Label Generation")
    st.markdown("✓ Gift Messages")
    st.markdown("✓ Label Merging")
    
    st.markdown("---")
    st.markdown('<div class="status-indicator"><div class="status-dot"></div><span>System Ready</span></div>', unsafe_allow_html=True)

# --------------------------------------
# MAIN CONTENT
# --------------------------------------
st.title("🪵 Charcuterie Board Order Manager")

st.markdown("""
**Professional board order processing & label generation system**  
Parse Amazon PDFs • Generate design-specific CSVs • Create labels • Merge shipments
""")

st.markdown("---")

# File Upload Section
st.markdown('<a id="upload-order"></a>', unsafe_allow_html=True)
st.markdown("## 📄 Upload Order")
uploaded = st.file_uploader(
    "Drop your Amazon packing slip PDF here",
    type=["pdf"],
    help="Upload the packing slip PDF from your Amazon board orders"
)

# --------------------------------------
# Parse PDF
# --------------------------------------
if uploaded:
    st.info("⏳ Reading and parsing your PDF...")

    all_pages = []
    with pdfplumber.open(uploaded) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            all_pages.append(text)

    records = []

    for page_text in all_pages:
        # Extract buyer name
        buyer_match = re.search(r"Ship To:\s*([\s\S]*?)Order ID:", page_text)
        buyer_name = ""
        if buyer_match:
            lines = [l.strip() for l in buyer_match.group(1).splitlines() if l.strip()]
            if lines:
                buyer_name = lines[0]

        # Extract order ID and date
        order_id = ""
        order_date = ""
        m_id = re.search(r"Order ID:\s*([\d\-]+)", page_text)
        if m_id:
            order_id = m_id.group(1).strip()
        m_date = re.search(r"Order Date:\s*([A-Za-z]{3,},?\s*[A-Za-z]+\s*\d{1,2},?\s*\d{4})", page_text)
        if m_date:
            order_date = m_date.group(1).strip()

        # Split into blocks by Customizations
        blocks = re.split(r"(?=Customizations:)", page_text)
        
        for block in blocks:
            if "Customizations:" not in block:
                continue

            # Extract quantity
            qty_match = re.search(r"Quantity\s*\n\s*(\d+)", block)
            quantity = qty_match.group(1) if qty_match else "1"

            # Extract design number
            design_number = extract_design_number(block)

            # Extract engraving type
            engraving_type = extract_engraving_type(block)

            # Extract board customization note
            board_note_match = re.search(r"Board Customization Note:\s*([^\n]+)", block)
            board_customization = clean_text(board_note_match.group(1)) if board_note_match else ""

            # Extract utensil letter
            utensil_match = re.search(r"Engraving Letter for (?:Cheese Knife Handles|Utensils):\s*([A-Z])", block, re.IGNORECASE)
            utensil_letter = utensil_match.group(1) if utensil_match else ""

            # Check for gift note
            gift_note = "YES" if re.search(r"Gift Note & Gift Bag:\s*Yes", block, re.IGNORECASE) else "NO"

            # Extract gift message
            gift_msg_match = re.search(
                r"Gift Card Note:\s*([\s\S]*?)(?=\n(?:Please CHECK|Grand total|Returning your item|Visit|Quantity|Order Totals|$))",
                block,
                re.IGNORECASE
            )
            gift_message = clean_text(gift_msg_match.group(1)) if gift_msg_match else ""

            records.append({
                "Order ID": order_id,
                "Order Date": order_date,
                "Buyer Name": buyer_name,
                "Quantity": quantity,
                "Design Number": design_number,
                "Engraving Type": engraving_type,
                "Board Customization Note": board_customization,
                "Manual Field 1": "",
                "Manual Field 2": "",
                "Manual Field 3": "",
                "Utensil Letter": utensil_letter,
                "Gift Note": gift_note,
                "Gift Message": gift_message
            })

    if not records:
        st.error("❌ No orders detected. Please check your PDF format.")
        st.stop()

    df = pd.DataFrame(records)
    df.index = df.index + 1
    
    st.success(f"✅ Successfully parsed {len(df)} line items from {df['Order ID'].nunique()} orders")
    
    with st.expander("📊 View Order Data"):
        st.dataframe(df, use_container_width=True)

    # --------------------------------------
    # Calculate Summary Statistics
    # --------------------------------------
    df['Quantity_Int'] = df['Quantity'].astype(int)
    
    total_boards = df['Quantity_Int'].sum()
    total_orders = df['Order ID'].nunique()
    gift_messages_needed = len(df[df['Gift Note'] == 'YES'])
    unique_designs = df['Design Number'].nunique()
    
    design_counts = df.groupby('Design Number')['Quantity_Int'].sum().sort_index()
    engraving_counts = df.groupby('Engraving Type')['Quantity_Int'].sum()

    # --------------------------------------
    # Dashboard Metrics
    # --------------------------------------
    st.markdown("---")
    st.markdown('<a id="dashboard"></a>', unsafe_allow_html=True)
    st.markdown("## 📊 Order Dashboard")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total Boards", total_boards)
    with col2:
        st.metric("Total Orders", total_orders)
    with col3:
        st.metric("Gift Messages", gift_messages_needed)
    with col4:
        st.metric("Unique Designs", unique_designs)
    
    # --------------------------------------
    # Design Breakdown
    # --------------------------------------
    st.markdown("---")
    st.markdown('<a id="design-breakdown"></a>', unsafe_allow_html=True)
    st.markdown("## 🎨 Design & Type Breakdown")
    
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.markdown("### 📐 Design Numbers")
        for design, count in design_counts.items():
            if design == "NO_DESIGN":
                st.markdown(f"**No Design / Blank Boards:** {count} boards")
            else:
                st.markdown(f"**Design {design}:** {count} boards")
    
    with col_right:
        st.markdown("### 🔨 Engraving Types")
        for eng_type, count in engraving_counts.items():
            st.markdown(f"**{eng_type}:** {count} boards")

    # --------------------------------------
    # Generate Files Section
    # --------------------------------------
    st.markdown("---")
    st.markdown('<a id="generate-files"></a>', unsafe_allow_html=True)
    st.markdown("## 📥 Generate & Download")
    
    if 'manufacturing_labels_buffer' not in st.session_state:
        st.session_state.manufacturing_labels_buffer = None
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("📦 Manufacturing Labels", use_container_width=True):
            with st.spinner("Generating manufacturing labels..."):
                pdf_data = generate_manufacturing_labels(df)
                st.session_state.manufacturing_labels_buffer = pdf_data
            st.success("✅ Labels generated!")
            st.download_button(
                label="⬇️ Download Manufacturing Labels",
                data=pdf_data,
                file_name="Board_Manufacturing_Labels.pdf",
                mime="application/pdf",
                use_container_width=True
            )
    
    with col2:
        gift_count = len(df[df['Gift Message'] != ""])
        if st.button(f"💌 Gift Messages ({gift_count})", use_container_width=True):
            with st.spinner("Generating gift message labels..."):
                pdf_data = generate_gift_message_labels(df)
            st.success("✅ Labels generated!")
            st.download_button(
                label="⬇️ Download Gift Message Labels",
                data=pdf_data,
                file_name="Gift_Message_Labels.pdf",
                mime="application/pdf",
                use_container_width=True
            )
    
    with col3:
        if st.button(f"📊 Design CSVs ({unique_designs})", use_container_width=True):
            with st.spinner("Generating design-specific CSV files..."):
                zip_data = generate_csv_files_by_design(df)
            st.success("✅ CSV files generated!")
            st.download_button(
                label="⬇️ Download All CSVs (ZIP)",
                data=zip_data,
                file_name="Board_Orders_by_Design.zip",
                mime="application/zip",
                use_container_width=True
            )

    # --------------------------------------
    # Label Merging Section
    # --------------------------------------
    st.markdown("---")
    st.markdown('<a id="label-merge"></a>', unsafe_allow_html=True)
    st.markdown("## 🔄 Merge Shipping & Manufacturing Labels")
    
    st.info("""
    **Instructions for Label Merging:**
    1. Generate Manufacturing Labels above (click the button)
    2. Upload your shipping labels PDF from Amazon/UPS
    3. Click merge to create a combined PDF
    """)
    
    shipping_labels_upload = st.file_uploader(
        "📤 Upload Shipping Labels PDF",
        type=["pdf"],
        key="shipping_labels",
        help="Upload the shipping labels PDF from Amazon or your carrier"
    )
    
    if shipping_labels_upload and st.session_state.manufacturing_labels_buffer:
        col_merge1, col_merge2 = st.columns([3, 1])
        
        with col_merge1:
            if st.button("🔀 Merge Labels Now", type="primary", use_container_width=True):
                with st.spinner("Merging shipping and manufacturing labels..."):
                    shipping_labels_upload.seek(0)
                    st.session_state.manufacturing_labels_buffer.seek(0)
                    
                    merged_pdf, num_shipping, num_manufacturing = merge_shipping_and_manufacturing_labels(
                        shipping_labels_upload,
                        st.session_state.manufacturing_labels_buffer,
                        df
                    )
                    
                    if merged_pdf:
                        st.success(f"✅ Successfully merged {num_shipping} shipping labels with {num_manufacturing} manufacturing labels!")
                        
                        multi_item_orders = df.groupby('Order ID').size()
                        multi_item_orders = multi_item_orders[multi_item_orders > 1]
                        
                        if len(multi_item_orders) > 0:
                            with st.expander(f"ℹ️ Found {len(multi_item_orders)} order(s) with multiple items"):
                                for order_id, count in multi_item_orders.items():
                                    buyer = df[df['Order ID'] == order_id]['Buyer Name'].iloc[0]
                                    st.write(f"• {buyer} ({order_id}): {count} boards")
                        
                        st.download_button(
                            label="⬇️ Download Merged Labels PDF",
                            data=merged_pdf,
                            file_name="Merged_Board_Labels.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
        
        with col_merge2:
            st.metric("Total Orders", df['Order ID'].nunique())
            st.metric("Total Items", len(df))
    
    elif shipping_labels_upload and not st.session_state.manufacturing_labels_buffer:
        st.warning("⚠️ Please generate Manufacturing Labels first (click the button above)")
    
    elif not shipping_labels_upload and st.session_state.manufacturing_labels_buffer:
        st.info("📤 Upload your shipping labels PDF above to enable merging")

    # --------------------------------------
    # NEW: Merge Labels by Design Section
    # --------------------------------------
    if shipping_labels_upload and st.session_state.manufacturing_labels_buffer:
        st.markdown("---")
        st.markdown('<a id="design-merge"></a>', unsafe_allow_html=True)
        st.markdown("## 🎨 Merge Labels by Design")
        
        st.info("""
        **Group and merge labels by design number:**
        - Each design gets its own merged PDF
        - Orders with mixed designs go into a separate PDF
        - Perfect for organizing production by design
        """)
        
        # Analyze orders to determine pure vs mixed designs
        order_design_map = {}
        for order_id in df['Order ID'].unique():
            order_items = df[df['Order ID'] == order_id]
            designs = order_items['Design Number'].unique().tolist()
            order_design_map[order_id] = designs
        
        # Count pure design orders per design
        design_order_counts = {}
        mixed_order_count = 0
        
        for order_id, designs in order_design_map.items():
            if len(designs) == 1:
                design_num = designs[0]
                design_order_counts[design_num] = design_order_counts.get(design_num, 0) + 1
            else:
                mixed_order_count += 1
        
        # Display buttons for each design
        if design_order_counts or mixed_order_count > 0:
            st.markdown("### 📦 Available Design Groups:")
            
            # Separate NO_DESIGN from numbered designs
            no_design_count = design_order_counts.pop("NO_DESIGN", 0)
            
            # Create columns for numbered design buttons (max 3 per row)
            designs_list = sorted([d for d in design_order_counts.keys() if d.isdigit()], key=lambda x: int(x))
            num_cols = 3
            
            for i in range(0, len(designs_list), num_cols):
                cols = st.columns(num_cols)
                for j, col in enumerate(cols):
                    if i + j < len(designs_list):
                        design_num = designs_list[i + j]
                        order_count = design_order_counts[design_num]
                        
                        with col:
                            button_key = f"design_{design_num}_merge"
                            if st.button(
                                f"📐 Design {design_num} ({order_count} orders)",
                                key=button_key,
                                use_container_width=True
                            ):
                                with st.spinner(f"Merging Design {design_num} labels..."):
                                    shipping_labels_upload.seek(0)
                                    st.session_state.manufacturing_labels_buffer.seek(0)
                                    
                                    merged_pdf, num_orders, num_items = merge_labels_by_design(
                                        shipping_labels_upload,
                                        st.session_state.manufacturing_labels_buffer,
                                        df,
                                        target_design=design_num,
                                        mixed_only=False
                                    )
                                    
                                    if merged_pdf:
                                        st.success(f"✅ Design {design_num}: Merged {num_orders} orders ({num_items} items)")
                                        st.download_button(
                                            label=f"⬇️ Download Design {design_num} Labels",
                                            data=merged_pdf,
                                            file_name=f"Design_{design_num}_Merged_Labels.pdf",
                                            mime="application/pdf",
                                            use_container_width=True,
                                            key=f"download_design_{design_num}"
                                        )
            
            # No Design / Blank Boards button
            if no_design_count > 0:
                st.markdown("### 🔲 Blank Boards:")
                if st.button(
                    f"🔲 No Design / Blank Boards ({no_design_count} orders)",
                    use_container_width=True,
                    key="no_design_merge"
                ):
                    with st.spinner("Merging blank board labels..."):
                        shipping_labels_upload.seek(0)
                        st.session_state.manufacturing_labels_buffer.seek(0)
                        
                        merged_pdf, num_orders, num_items = merge_labels_by_design(
                            shipping_labels_upload,
                            st.session_state.manufacturing_labels_buffer,
                            df,
                            target_design="NO_DESIGN",
                            mixed_only=False
                        )
                        
                        if merged_pdf:
                            st.success(f"✅ Blank Boards: Merged {num_orders} orders ({num_items} items)")
                            st.download_button(
                                label="⬇️ Download Blank Board Labels",
                                data=merged_pdf,
                                file_name="No_Design_Blank_Boards.pdf",
                                mime="application/pdf",
                                use_container_width=True,
                                key="download_no_design"
                            )
            
            # Mixed design orders button
            if mixed_order_count > 0:
                st.markdown("### 🔀 Mixed Design Orders:")
                if st.button(
                    f"🔀 Mixed Designs ({mixed_order_count} orders)",
                    use_container_width=True,
                    key="mixed_designs_merge"
                ):
                    with st.spinner("Merging mixed design labels..."):
                        shipping_labels_upload.seek(0)
                        st.session_state.manufacturing_labels_buffer.seek(0)
                        
                        merged_pdf, num_orders, num_items = merge_labels_by_design(
                            shipping_labels_upload,
                            st.session_state.manufacturing_labels_buffer,
                            df,
                            target_design=None,
                            mixed_only=True
                        )
                        
                        if merged_pdf:
                            st.success(f"✅ Mixed Designs: Merged {num_orders} orders ({num_items} items)")
                            
                            # Show which orders are mixed
                            with st.expander("ℹ️ View Mixed Design Orders"):
                                for order_id, designs in order_design_map.items():
                                    if len(designs) > 1:
                                        buyer = df[df['Order ID'] == order_id]['Buyer Name'].iloc[0]
                                        designs_str = ", ".join([f"Design {d}" for d in designs])
                                        st.write(f"• {buyer} ({order_id}): {designs_str}")
                            
                            st.download_button(
                                label="⬇️ Download Mixed Design Labels",
                                data=merged_pdf,
                                file_name="Mixed_Design_Orders.pdf",
                                mime="application/pdf",
                                use_container_width=True,
                                key="download_mixed_designs"
                            )

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #a0aec0; padding: 20px;'>
    <p><strong>Charcuterie Board Order Manager v1.1 Dark</strong></p>
    <p>Professional board order processing & label generation system</p>
</div>
""", unsafe_allow_html=True)
