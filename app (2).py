import os
import uuid
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st
from PIL import Image
import smtplib
import ssl
from email.message import EmailMessage

DATA_DIR = "data"
USERS_FILE = os.path.join(DATA_DIR, "users.xlsx")
ORDERS_FILE = os.path.join(DATA_DIR, "orders.xlsx")
ADMIN_EMAIL = "admin@adedgeos.com"
ADMIN_PASSWORD = "admin123"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465
EMAIL_SENDER = "adedgeosolution@gmail.com"
EMAIL_PASSWORD = "Aded@2025geo"
EMAIL_SUPPORT = "adedgeosolution@gmail.com"

st.set_page_config(
    page_title="Infer Sustainability | ADED Geos - Geospatial Data Portal",
    page_icon="🛰️",
    layout="wide",
    initial_sidebar_state="expanded",
)


def init_data():
    os.makedirs(DATA_DIR, exist_ok=True)

    if not os.path.exists(USERS_FILE):
        pd.DataFrame(
            columns=[
                "full_name",
                "contact_number",
                "whatsapp_number",
                "email",
                "dob",
                "occupation",
                "designation",
                "organisation",
                "aadhar_path",
                "approval_status",
                "registered_at",
            ]
        ).to_excel(USERS_FILE, index=False)

    if not os.path.exists(ORDERS_FILE):
        pd.DataFrame(
            columns=[
                "order_id",
                "user_email",
                "product_type",
                "category",
                "delivery_days",
                "expected_delivery_date",
                "product_description",
                "additional_description",
                "schedule_meeting",
                "meeting_date_1",
                "meeting_time_1",
                "meeting_date_2",
                "meeting_time_2",
                "meeting_date_3",
                "meeting_time_3",
                "additional_note",
                "urgent",
                "status",
                "created_at",
            ]
        ).to_excel(ORDERS_FILE, index=False)


@st.cache_data(show_spinner=False)
def load_users():
    users = pd.read_excel(USERS_FILE)
    if "approved" in users.columns and "approval_status" not in users.columns:
        users = users.copy()
        users["approval_status"] = users["approved"].map({True: "Approved", False: "Pending"})
        users = users.drop(columns=["approved"])
    if "approval_status" not in users.columns:
        users = users.copy()
        users["approval_status"] = "Pending"
    return users


@st.cache_data(show_spinner=False)
def load_orders():
    return pd.read_excel(ORDERS_FILE)


def save_users(df):
    df.to_excel(USERS_FILE, index=False)
    st.cache_data.clear()


def save_orders(df):
    df.to_excel(ORDERS_FILE, index=False)
    st.cache_data.clear()


def send_email(to_address: str, subject: str, body: str) -> bool:
    try:
        message = EmailMessage()
        message["From"] = EMAIL_SENDER
        message["To"] = to_address
        message["Subject"] = subject
        message.set_content(body)

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as smtp:
            smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
            smtp.send_message(message)
        return True
    except Exception as exc:
        st.warning(f"Email send failed: {exc}")
        return False


def get_user_by_email(email):
    users = load_users()
    result = users[users["email"] == email]
    return result.iloc[0] if not result.empty else None


def show_banner():
    st.markdown("""
    <div style='display:flex; align-items:center; gap:20px; padding:24px; border-radius:20px; background: linear-gradient(135deg, #1e3c72 0%, #2a5298 50%, #7e8ba3 100%); color: white; box-shadow: 0 8px 32px rgba(0,0,0,0.2);'>
        <div style='width:100px; height:100px; border-radius:20px; background: white; display:flex; align-items:center; justify-content:center; box-shadow: 0 4px 15px rgba(0,0,0,0.1);'>
            <img src='file:///workspaces/InferSust-Cart/InferLogo.png' style='width:90%; height:90%; object-fit:contain;'/>
        </div>
        <div>
            <p style='margin:0; font-size:14px; letter-spacing:2px; font-weight:600; text-transform:uppercase; color:#e8f0f7;'>Infer Sustainability | ADED Geos</p>
            <h1 style='margin:8px 0 12px; font-size:36px; font-weight:800; letter-spacing:-0.5px;'>Order your data, 'cause you're Spatial to us.</h1>
            <p style='margin:0; font-size:15px; max-width:800px; line-height:1.6; color:#e8f0f7;'>An interactive geospatial dashboard for online data orders, approvals, and delivery tracking — agriculture, climate, environment, forest, road network, mining and more.</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div style='margin-top:16px; padding:16px; background:linear-gradient(90deg, #f0f4f8 0%, #e3f2fd 100%); border-radius:12px; border-left:4px solid #2a5298;'>
        <div style='display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:24px;'>
            <div>
                <p style='margin:0; font-size:13px; color:#666; font-weight:600; text-transform:uppercase; letter-spacing:1px;'>📧 Support Email</p>
                <p style='margin:4px 0 0; font-size:16px; font-weight:600; color:#2a5298;'>adedgeosolution@gmail.com</p>
            </div>
            <div>
                <p style='margin:0; font-size:13px; color:#666; font-weight:600; text-transform:uppercase; letter-spacing:1px;'>📱 Contact Number</p>
                <p style='margin:4px 0 0; font-size:16px; font-weight:600; color:#2a5298;'>+91 9054084334</p>
            </div>
            <div>
                <p style='margin:0; font-size:13px; color:#666; font-weight:600; text-transform:uppercase; letter-spacing:1px;'>⏱️ Delivery & Payment</p>
                <p style='margin:4px 0 0; font-size:16px; font-weight:600; color:#2a5298;'>Shared via Email</p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def register_user():
    st.header("📝 Register to request geospatial data")
    with st.form("registration_form", clear_on_submit=False):
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### Personal Information")
            full_name = st.text_input("Full Name", max_chars=80)
            contact_number = st.text_input("Contact Number")
            whatsapp_number = st.text_input("WhatsApp Number")
            email = st.text_input("Official Email ID")
            dob = st.date_input("Date of Birth", max_value=date.today())
        with col2:
            st.markdown("### Professional Information")
            occupation = st.text_input("Occupation")
            designation = st.text_input("Designation")
            organisation = st.text_input("Organisation")
            aadhar_file = st.file_uploader("Upload Aadhaar / identity image", type=["png", "jpg", "jpeg", "pdf"])
            agree = st.checkbox("✅ I confirm the above details are correct.")

        submitted = st.form_submit_button("Submit registration request", use_container_width=True)

        if submitted:
            if not all([full_name, contact_number, whatsapp_number, email, occupation, designation, organisation, aadhar_file, agree]):
                st.error("Please complete all fields and upload the Aadhaar image to register.")
                return

            users = load_users()
            if email in users["email"].astype(str).values:
                st.warning("A registration request already exists for this email. Please wait for admin approval or contact support.")
                return

            aadhar_path = os.path.join(DATA_DIR, f"aadhar_{uuid.uuid4().hex}{os.path.splitext(aadhar_file.name)[1]}")
            with open(aadhar_path, "wb") as f:
                f.write(aadhar_file.getbuffer())

            new_user = {
                "full_name": full_name,
                "contact_number": contact_number,
                "whatsapp_number": whatsapp_number,
                "email": email,
                "dob": dob.strftime("%Y-%m-%d"),
                "occupation": occupation,
                "designation": designation,
                "organisation": organisation,
                "aadhar_path": aadhar_path,
                "approval_status": "Pending",
                "registered_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            users = pd.concat([users, pd.DataFrame([new_user])], ignore_index=True)
            save_users(users)

            user_subject = "AdEdgeOS registration request received"
            user_body = f"Dear {full_name},\n\nYour registration request has been received and is pending admin approval. We will notify you once the account is approved.\n\nThank you,\nAdEdgeOS Team"
            send_email(email, user_subject, user_body)

            admin_subject = "New registration request submitted"
            admin_body = f"A new registration request has been submitted:\n\nName: {full_name}\nEmail: {email}\nContact: {contact_number}\nOrganisation: {organisation}\nDesignation: {designation}\n\nPlease review and approve or reject via the admin dashboard."
            send_email(EMAIL_SUPPORT, admin_subject, admin_body)

            st.success("✅ Registration request submitted successfully. Admin approval is required before you can place an order.")
            st.info("Please use the login section after approval to order geospatial data.")


def admin_dashboard():
    st.header("🔧 Admin Portal")
    users = load_users()
    orders = load_orders()

    st.markdown("### 📊 Dashboard Summary")
    pending_orders = len(orders[orders["status"] == "Pending"])
    delivered_orders = len(orders[orders["status"] == "Delivered"])
    last_month_cutoff = datetime.now() - timedelta(days=30)
    last_month_orders = len(orders[pd.to_datetime(orders["created_at"], errors="coerce") >= last_month_cutoff])
    total_orders = len(orders)
    progress = int((delivered_orders / max(total_orders, 1)) * 100)

    total_users = len(users)
    pending_users_count = len(users[users["approval_status"] == "Pending"])
    approved_users_count = len(users[users["approval_status"] == "Approved"])
    rejected_users_count = len(users[users["approval_status"] == "Rejected"])

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("⏳ Pending Orders", pending_orders)
    col2.metric("✅ Delivered Orders", delivered_orders)
    col3.metric("📅 Orders (30 days)", last_month_orders)
    col4.metric("🎯 Delivery Progress", f"{progress}%")

    st.markdown("---")
    st.markdown("### 👥 User Approval Summary")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("👤 Total Users", total_users)
    col2.metric("⏳ Pending Approvals", pending_users_count)
    col3.metric("✅ Approved Users", approved_users_count)
    col4.metric("❌ Rejected Users", rejected_users_count)

    st.markdown("---")
    st.markdown("### 📋 User Approval Queue")
    pending_users = users[users["approval_status"] == "Pending"]
    if pending_users.empty:
        st.info("✅ No pending user approvals.")
    else:
        st.dataframe(pending_users.drop(columns=["aadhar_path"]), use_container_width=True)
        selected_email = st.selectbox("Select a user to approve", options=pending_users["email"].tolist())
        if selected_email:
            col1, col2 = st.columns(2)
            with col1:
                if st.button("✅ Approve selected user", use_container_width=True):
                    users.loc[users["email"] == selected_email, "approval_status"] = "Approved"
                    save_users(users)
                    approved_user = get_user_by_email(selected_email)
                    if approved_user is not None:
                        subject = "Your AdEdgeOS account has been approved"
                        body = f"Dear {approved_user['full_name']},\n\nYour AdEdgeOS account has been approved. You can now login and place an order.\n\nThank you,\nAdEdgeOS Team"
                        send_email(approved_user["email"], subject, body)
                    st.success("✅ User approved. The user may now login and place an order.")
                    st.rerun()
            with col2:
                if st.button("❌ Reject selected user", use_container_width=True):
                    users.loc[users["email"] == selected_email, "approval_status"] = "Rejected"
                    save_users(users)
                    rejected_user = get_user_by_email(selected_email)
                    if rejected_user is not None:
                        subject = "Your AdEdgeOS account request was rejected"
                        body = f"Dear {rejected_user['full_name']},\n\nYour registration request has been rejected by the admin. Please contact support at {EMAIL_SUPPORT} for additional details.\n\nThank you,\nAdEdgeOS Team"
                        send_email(rejected_user["email"], subject, body)
                    st.success("❌ User request has been rejected.")
                    st.rerun()

    st.markdown("---")
    st.markdown("### 📦 Order Management")
    if orders.empty:
        st.info("No orders submitted yet.")
    else:
        st.dataframe(orders, use_container_width=True)
        selected_order = st.selectbox("Select order to update", options=orders["order_id"].tolist())
        if selected_order:
            order_row = orders[orders["order_id"] == selected_order].iloc[0]
            st.write("**Selected order details**")
            st.write(order_row.to_dict())

            col1, col2 = st.columns(2)
            with col1:
                new_status = st.selectbox("Update status", ["Pending", "In progress", "Delivered", "Cancelled"], index=["Pending", "In progress", "Delivered", "Cancelled"].index(order_row["status"]) if order_row["status"] in ["Pending", "In progress", "Delivered", "Cancelled"] else 0)
            with col2:
                new_note = st.text_area("Admin note", value=order_row.get("additional_note", ""))

            col1, col2, col3 = st.columns(3)
            with col1:
                if st.button("💾 Save order update", use_container_width=True):
                    orders.loc[orders["order_id"] == selected_order, "status"] = new_status
                    orders.loc[orders["order_id"] == selected_order, "additional_note"] = new_note
                    save_orders(orders)
                    selected = orders[orders["order_id"] == selected_order].iloc[0]
                    customer = get_user_by_email(selected["user_email"])
                    if customer is not None:
                        subject = f"Order {selected_order} status updated"
                        body = f"Dear {customer['full_name']},\n\nYour order {selected_order} has been updated to '{new_status}'.\nAdmin note: {new_note}\n\nThank you,\nAdEdgeOS Team"
                        send_email(customer["email"], subject, body)
                    st.success("✅ Order updated successfully.")
                    st.rerun()
            with col2:
                if st.button("🗑️ Delete selected order", use_container_width=True):
                    orders = orders[orders["order_id"] != selected_order]
                    save_orders(orders)
                    st.success("🗑️ Order removed from the system.")
                    st.rerun()

    st.markdown("---")
    st.markdown("### 👥 User Dataset Management")
    st.dataframe(users.drop(columns=["aadhar_path"]), use_container_width=True)
    selected_user = st.selectbox("Select a user to remove", options=users["email"].tolist())
    if selected_user and st.button("🗑️ Delete selected user record", use_container_width=True):
        users = users[users["email"] != selected_user]
        save_users(users)
        st.success("✅ User record deleted.")
        st.rerun()


def user_portal(user):
    st.header(f"Welcome, {user['full_name']} 👋")
    if user["approval_status"] != "Approved":
        if user["approval_status"] == "Rejected":
            st.error("❌ Your registration request has been rejected. Please contact the administrator for further details.")
        else:
            st.warning("⏳ Your account is pending admin approval. You will receive an email once the admin approves your registration.")
        return

    st.success("✅ Your account has been approved. Submit a geospatial product request below.")
    
    with st.form("order_form"):
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### 📋 Order Details")
            product_type = st.selectbox("Select data format", ["Map", "Graph", "Table"])
            category = st.selectbox("Data category", ["Agriculture", "Climate", "Environment", "Forest", "Road Network", "Mining", "Others"])
            delivery_days = st.selectbox("Expected delivery window", ["3 days", "5 days", "7 days", "14 days"])
            product_description = st.text_area("Product Description (Requirement)", height=120, placeholder="Describe your geospatial data requirements in detail...")
        with col2:
            st.markdown("### 📝 Additional Information")
            additional_description = st.text_area("Additional Description (Requirement)", height=120, placeholder="Any additional specifications or details...")
            urgent = st.checkbox("🚀 Mark order as urgent")
            schedule_meeting = st.checkbox("📅 Schedule a meeting with our team")

        if schedule_meeting:
            st.markdown("---")
            st.markdown("""
            <div style='padding:16px; background:linear-gradient(135deg, #e3f2fd 0%, #f3e5f5 100%); border-radius:12px; border-left:4px solid #7c4dff;'>
                <h4 style='margin:0 0 16px; color:#4527a0; display:flex; align-items:center; gap:8px;'>📅 Meeting Preferences</h4>
            </div>
            """, unsafe_allow_html=True)
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown("**Option 1**")
                meeting_date_1 = st.date_input("Preferable meeting date", min_value=date.today(), key="date1")
                meeting_time_1 = st.time_input("Time range", key="time1")
            with col2:
                st.markdown("**Option 2**")
                meeting_date_2 = st.date_input("Preferable meeting date", min_value=date.today(), key="date2")
                meeting_time_2 = st.time_input("Time range", key="time2")
            with col3:
                st.markdown("**Option 3**")
                meeting_date_3 = st.date_input("Preferable meeting date", min_value=date.today(), key="date3")
                meeting_time_3 = st.time_input("Time range", key="time3")
        else:
            meeting_date_1 = meeting_time_1 = meeting_date_2 = meeting_time_2 = meeting_date_3 = meeting_time_3 = ""

        st.markdown("---")
        additional_note = st.text_area("Additional Note", height=80, help="Any specific requirements or notes for our team", placeholder="Meeting schedule and payment delivery details will be shared by email.")

        st.markdown("""
        <div style='padding:14px; background:#f5f5f5; border-radius:10px; margin-top:16px;'>
            <p style='margin:0; font-size:13px; color:#424242; line-height:1.6;'>
                <strong>ℹ️ Important Information:</strong><br>
                • Meeting schedule will be communicated via email<br>
                • Payment options and delivery dates will be shared via email<br>
                • Contact us at adedgeosolution@gmail.com or +91 9054084334 for any queries
            </p>
        </div>
        """, unsafe_allow_html=True)

        if st.form_submit_button("🚀 Submit order request", use_container_width=True):
            if not product_description:
                st.error("Please add a product description to continue.")
            else:
                orders = load_orders()
                order_id = f"AEO-{uuid.uuid4().hex[:8].upper()}"
                new_order = {
                    "order_id": order_id,
                    "user_email": user["email"],
                    "product_type": product_type,
                    "category": category,
                    "delivery_days": delivery_days,
                    "expected_delivery_date": (date.today() + timedelta(days=int(delivery_days.split()[0]))).strftime("%Y-%m-%d"),
                    "product_description": product_description,
                    "additional_description": additional_description,
                    "schedule_meeting": schedule_meeting,
                    "meeting_date_1": meeting_date_1 if schedule_meeting else "",
                    "meeting_time_1": meeting_time_1.strftime("%H:%M") if schedule_meeting else "",
                    "meeting_date_2": meeting_date_2 if schedule_meeting else "",
                    "meeting_time_2": meeting_time_2.strftime("%H:%M") if schedule_meeting else "",
                    "meeting_date_3": meeting_date_3 if schedule_meeting else "",
                    "meeting_time_3": meeting_time_3.strftime("%H:%M") if schedule_meeting else "",
                    "additional_note": additional_note,
                    "urgent": urgent,
                    "status": "Pending",
                    "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
                orders = pd.concat([orders, pd.DataFrame([new_order])], ignore_index=True)
                save_orders(orders)
                order_subject = "AdEdgeOS order request received"
                order_body = f"Dear {user['full_name']},\n\nYour order request {order_id} has been submitted successfully.\nCategory: {category}\nFormat: {product_type}\nExpected delivery date: {new_order['expected_delivery_date']}\n\nPayment and delivery details will be shared by email.\n\nThank you,\nAdEdgeOS Team"
                send_email(user["email"], order_subject, order_body)
                admin_order_subject = "New AdEdgeOS order request submitted"
                admin_order_body = f"A new order request has been placed by {user['full_name']} ({user['email']}).\nOrder ID: {order_id}\nCategory: {category}\nFormat: {product_type}\nExpected delivery date: {new_order['expected_delivery_date']}\n\nPlease follow up with the customer as needed."
                send_email(EMAIL_SUPPORT, admin_order_subject, admin_order_body)
                st.success("✅ Request submitted. You will get delivery and payment details by mail soon.")

    st.markdown("---")
    st.subheader("📦 Your order history")
    orders = load_orders()
    user_orders = orders[orders["user_email"] == user["email"]]
    if user_orders.empty:
        st.info("You have not placed any orders yet.")
    else:
        st.dataframe(user_orders.drop(columns=["user_email"]), use_container_width=True)


def login_portal():
    st.sidebar.title("🔐 Access Dashboard")
    access_type = st.sidebar.radio("Choose portal", ["📝 Register", "👤 User Login", "🔧 Admin Login"])

    if access_type == "📝 Register":
        register_user()
    elif access_type == "🔧 Admin Login":
        st.sidebar.subheader("Admin access")
        admin_email = st.sidebar.text_input("Admin email", value=ADMIN_EMAIL)
        admin_password = st.sidebar.text_input("Password", type="password")
        if st.sidebar.button("🔓 Login as admin", use_container_width=True):
            if admin_email == ADMIN_EMAIL and admin_password == ADMIN_PASSWORD:
                st.session_state["role"] = "admin"
                st.rerun()
            else:
                st.sidebar.error("❌ Invalid admin credentials.")
    else:
        st.sidebar.subheader("User login")
        email = st.sidebar.text_input("Official Email ID")
        contact = st.sidebar.text_input("Contact Number")
        if st.sidebar.button("🔓 Login as user", use_container_width=True):
            user = get_user_by_email(email)
            if user is None:
                st.sidebar.error("❌ No registration found for this email. Please register first.")
            elif str(user["contact_number"]).strip() != str(contact).strip():
                st.sidebar.error("❌ Contact number does not match the registered account.")
            else:
                st.session_state["role"] = "user"
                st.session_state["user_email"] = email
                st.rerun()

    if st.session_state.get("role") == "admin":
        admin_dashboard()
    elif st.session_state.get("role") == "user":
        user = get_user_by_email(st.session_state.get("user_email", ""))
        if user is not None:
            user_portal(user)
        else:
            st.error("❌ Unable to find user session. Please log in again.")

    if st.sidebar.button("🚪 Logout", use_container_width=True):
        st.session_state.clear()
        st.rerun()


def main():
    # Custom CSS for enhanced appearance
    st.markdown("""
    <style>
    /* General styling */
    :root {
        --primary-color: #2a5298;
        --secondary-color: #7c4dff;
        --success-color: #4caf50;
        --warning-color: #ff9800;
        --danger-color: #f44336;
    }
    
    /* Button styling */
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        padding: 10px 24px;
        transition: all 0.3s ease;
        border: none;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
    
    /* Form styling */
    .stForm {
        border-radius: 12px;
        padding: 20px;
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    
    /* Metric styling */
    .stMetricValue {
        font-size: 28px;
        font-weight: 700;
        color: #2a5298;
    }
    
    /* Sidebar styling */
    .stSidebar {
        background: linear-gradient(180deg, #f5f7fa 0%, #e8f0f7 100%);
    }
    
    /* Heading styling */
    h1, h2, h3, h4, h5, h6 {
        color: #1e3c72;
        font-weight: 700;
    }
    
    /* Info/Success/Warning/Error boxes */
    .stSuccess, .stInfo, .stWarning, .stError {
        border-radius: 10px;
        padding: 15px 20px;
        font-weight: 500;
    }
    </style>
    """, unsafe_allow_html=True)
    
    init_data()
    show_banner()
    login_portal()


if __name__ == "__main__":
    main()
