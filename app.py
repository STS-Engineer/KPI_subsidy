import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import psycopg2
from psycopg2 import pool
from flask import Flask, request
from apscheduler.schedulers.background import BackgroundScheduler
import pytz
import traceback
from urllib.parse import quote_plus
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

app = Flask(__name__)

# ---------- Configuration ----------
PORT = int(os.getenv('PORT', 5000))

# ---------- PostgreSQL Connection Pool ----------
db_pool = pool.SimpleConnectionPool(
    1, 20,
    user=os.getenv('DB_USER', "administrationSTS"),
    host=os.getenv('DB_HOST', "avo-adb-002.postgres.database.azure.com"),
    database=os.getenv('DB_NAME', "Subsidy_DB"),
    password=os.getenv('DB_PASSWORD', "St$@0987"),
    port=int(os.getenv('DB_PORT', 5432)),
    sslmode="require"
)

# ---------- Email Configuration ----------
SMTP_SERVER = os.getenv('EMAIL_HOST', "smtp.office365.com")
SMTP_PORT = int(os.getenv('EMAIL_PORT', 587))
EMAIL_USER = os.getenv('EMAIL_USER', "administration.STS@avocarbon.com")
EMAIL_PASSWORD = os.getenv('EMAIL_PASS', "shnlgdyfbcztbhxn")

# ========================================
# HELPER FUNCTIONS
# ========================================

def _base_url():
    """Get base URL for email links"""
    try:
        if request and request.host_url:
            return request.host_url.rstrip('/')
    except RuntimeError:
        pass
    return os.getenv('APP_BASE_URL', f'http://localhost:{PORT}')

def get_current_iso_week():
    """Get current ISO week in format YYYY-Wxx (e.g., 2025-W43)"""
    now = datetime.now()
    iso_calendar = now.isocalendar()
    return f"{iso_calendar[0]}-W{iso_calendar[1]:02d}"

def get_responsible_with_kpis(responsible_id, week):
    """Fetch responsible info and their KPIs for a given week"""
    conn = db_pool.getconn()
    try:
        with conn.cursor() as cur:
            # Fetch responsible info
            cur.execute(
                """
                SELECT responsible_id, name, email
                FROM public."Responsible"
                WHERE responsible_id = %s
                """,
                (responsible_id,),
            )
            responsible = cur.fetchone()
            if not responsible:
                raise Exception("Responsible not found")

            # Fetch KPIs
            cur.execute(
                """
                SELECT kv.kpi_values_id, kv.value, kv.week, kv.analyse, kv.actions_correctives,
                       k.kpi_id, k."KPI_name", k."KPI_objectif"
                FROM public.kpi_values kv
                JOIN public."Kpi" k ON kv.kpi_id = k.kpi_id
                WHERE kv.responsible_id = %s AND kv.week = %s
                ORDER BY k.kpi_id ASC
                """,
                (responsible_id, week),
            )
            kpis = cur.fetchall()

            return {
                'responsible': {
                    'responsible_id': responsible[0],
                    'name': responsible[1],
                    'email': responsible[2],
                },
                'kpis': [
                    {
                        'kpi_values_id': kpi[0],
                        'value': kpi[1],
                        'week': kpi[2],
                        'analyse': kpi[3],
                        'actions_correctives': kpi[4],
                        'kpi_id': kpi[5],
                        'KPI_name': kpi[6],
                        'KPI_objectif': kpi[7],
                    }
                    for kpi in kpis
                ],
            }
    except Exception as e:
        print(f"❌ Database error in get_responsible_with_kpis: {str(e)}")
        raise
    finally:
        db_pool.putconn(conn)

def send_kpi_email(responsible_id, responsible_name, responsible_email, kpi_name, week):
    """Send KPI email with a link to the form for a specific responsible"""
    try:
        base = _base_url()
        form_link = f"{base}/form?responsible_id={responsible_id}&week={quote_plus(week)}"

        msg = MIMEMultipart()
        msg['From'] = f'"Administration STS" <{EMAIL_USER}>'
        msg['To'] = responsible_email
        msg['Subject'] = f"KPI Report - {kpi_name} - Week {week}"

        html_content = f"""
        <!DOCTYPE html>
        <html>
        <body style="font-family:Arial,sans-serif; background:#f7f7f7; padding:24px;">
          <div style="max-width:640px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;border:1px solid #eee">
            <div style="background:#0078D7;color:#fff;padding:20px 24px">
              <h2 style="margin:0;font-weight:600">KPI Report – {responsible_name}</h2>
              <div style="margin-top:8px;font-size:14px">Week {week}</div>
            </div>
            <div style="padding:24px">
              <p>Hello {responsible_name},</p>
              <p>The KPI <strong>{kpi_name}</strong> is due for reporting for week <strong>{week}</strong>.</p>
              <p>Please click the link below to fill out your KPI analysis and corrective actions:</p>
              <p style="text-align:center;margin:28px 0">
                <a href="{form_link}"
                   style="display:inline-block;padding:12px 20px;border-radius:6px;background:#0078D7;color:#fff;text-decoration:none;font-weight:600">
                  Open KPI Form
                </a>
              </p>
              <p style="font-size:12px;color:#666;margin-top:24px">
                This is an automated reminder from the KPI tracking system.
              </p>
            </div>
          </div>
        </body>
        </html>
        """
        msg.attach(MIMEText(html_content, 'html'))

        # Use authenticated SMTP with TLS encryption
        print(f"📧 Connecting to {SMTP_SERVER}:{SMTP_PORT}...")
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=10) as server:
            server.starttls()  # Enable TLS encryption
            print(f"🔐 Starting TLS encryption...")
            server.login(EMAIL_USER, EMAIL_PASSWORD)  # Authenticate with credentials
            print(f"✅ Authenticated as {EMAIL_USER}")
            server.send_message(msg)

        print(f"✅ Email sent successfully to {responsible_email} for KPI: {kpi_name}")
        return True

    except Exception as e:
        print(f"❌ Failed to send email to {responsible_email}: {str(e)}")
        traceback.print_exc()
        return False

# ========================================
# SCHEDULER FUNCTIONS
# ========================================

def get_due_kpis_with_responsibles():
    """
    Fetch all KPIs that are due (frequence_de_envoi <= NOW) 
    along with their assigned responsibles for the current week.
    Returns: List of tuples (kpi_id, kpi_name, responsible_id, resp_name, email, week)
    """
    conn = db_pool.getconn()
    current_week = get_current_iso_week()

    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT DISTINCT 
                    k.kpi_id,
                    k."KPI_name",
                    r.responsible_id,
                    r.name,
                    r.email,
                    kv.week
                FROM public."Kpi" k
                JOIN public.kpi_values kv ON kv.kpi_id = k.kpi_id
                JOIN public."Responsible" r ON r.responsible_id = kv.responsible_id
                WHERE k.frequence_de_envoi <= NOW()
                  AND kv.week = %s
                ORDER BY k.kpi_id, r.responsible_id
                """,
                (current_week,),
            )

            results = cur.fetchall()
            print(f"📊 Found {len(results)} KPI-Responsible combinations due for week {current_week}")
            return results

    except Exception as e:
        print(f"❌ Error fetching due KPIs: {str(e)}")
        traceback.print_exc()
        return []
    finally:
        db_pool.putconn(conn)

def update_kpi_created_at(kpi_id):
    """
    Update created_at = NOW() to trigger recalculation of frequence_de_envoi
    This will automatically calculate the next send date based on the frequency rule
    """
    conn = db_pool.getconn()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                UPDATE public."Kpi"
                SET created_at = NOW()
                WHERE kpi_id = %s
                RETURNING kpi_id, "KPI_name", created_at, frequence_de_envoi
                """,
                (kpi_id,),
            )

            result = cur.fetchone()
            conn.commit()

            if result:
                kpi_id, kpi_name, new_created, new_freq = result
                print(f"   ✅ Updated '{kpi_name}' (ID:{kpi_id})")
                print(f"      New created_at: {new_created}")
                print(f"      Next send scheduled: {new_freq}")
                return True
            else:
                print(f"   ⚠️ KPI {kpi_id} not found for update")
                return False

    except Exception as e:
        print(f"   ❌ Error updating KPI {kpi_id}: {str(e)}")
        traceback.print_exc()
        conn.rollback()
        return False
    finally:
        db_pool.putconn(conn)

def scheduled_email_task():
    """
    Main automated scheduler task:
    1. Gets current ISO week
    2. Finds all KPIs where frequence_de_envoi <= NOW()
    3. Sends email to each responsible for their due KPIs
    4. Updates created_at to trigger next cycle calculation
    """
    print(f"\n{'='*70}")
    print(f"⏰ SCHEDULED TASK RUNNING at {datetime.now()}")
    print(f"{'='*70}")

    current_week = get_current_iso_week()
    print(f"📅 Current ISO Week: {current_week}")

    # Get all due KPIs with their responsibles
    due_records = get_due_kpis_with_responsibles()

    if not due_records:
        print("ℹ️  No KPIs are due for sending at this time.")
        print(f"{'='*70}\n")
        return

    kpis_processed = set()
    emails_sent = 0
    emails_failed = 0

    print(f"\n📧 Processing {len(due_records)} KPI-Responsible combination(s):\n")

    # Send email to each responsible for their due KPIs
    for kpi_id, kpi_name, responsible_id, resp_name, email, week in due_records:
        print(f"📤 Sending KPI reminder:")
        print(f"   KPI: {kpi_name} (ID: {kpi_id})")
        print(f"   To: {resp_name} ({email})")
        print(f"   Week: {week}")

        try:
            success = send_kpi_email(responsible_id, resp_name, email, kpi_name, week)

            if success:
                emails_sent += 1
                kpis_processed.add(kpi_id)
                print(f"   ✅ Email sent successfully\n")
            else:
                emails_failed += 1
                print(f"   ❌ Email failed to send\n")

        except Exception as e:
            emails_failed += 1
            print(f"   ❌ Exception: {str(e)}\n")
            traceback.print_exc()

    # Update KPIs to schedule next send
    if kpis_processed:
        print(f"\n🔄 Updating {len(kpis_processed)} KPI(s) for next cycle:\n")

        for kpi_id in kpis_processed:
            update_kpi_created_at(kpi_id)

    print(f"\n{'='*70}")
    print(f"✅ TASK COMPLETED:")
    print(f"   📧 Emails sent: {emails_sent}")
    print(f"   ❌ Emails failed: {emails_failed}")
    print(f"   🔄 KPIs updated: {len(kpis_processed)}")
    print(f"{'='*70}\n")

# ========================================
# FLASK ROUTES
# ========================================

@app.route('/')
def home():
    """Home page with system status"""
    next_run = scheduler.get_jobs()[0].next_run_time if scheduler.get_jobs() else "Not scheduled"
    
    return f'''
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>KPI Automation System</title>
        <style>
            body {{ font-family: 'Segoe UI', sans-serif; background: #f4f6f9; padding: 40px; }}
            .container {{ max-width: 800px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
            h1 {{ color: #0078D7; }}
            .status {{ background: #e7f3ff; padding: 15px; border-radius: 5px; margin: 20px 0; border-left: 4px solid #0078D7; }}
            .info {{ margin: 10px 0; }}
            .label {{ font-weight: 600; color: #333; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>🎯 KPI Automation System</h1>
            <div class="status">
                <div class="info"><span class="label">Status:</span> ✅ Running</div>
                <div class="info"><span class="label">Current Week:</span> {get_current_iso_week()}</div>
                <div class="info"><span class="label">Next Scheduled Check:</span> {next_run}</div>
                <div class="info"><span class="label">Server Time:</span> {datetime.now()}</div>
                <div class="info"><span class="label">Email Server:</span> {SMTP_SERVER}:{SMTP_PORT}</div>
            </div>
            <p>The system automatically checks for due KPIs and sends email notifications to responsible parties based on the <code>frequence_de_envoi</code> schedule.</p>
        </div>
    </body>
    </html>
    '''

@app.route('/form')
def form_page():
    """Display KPI form"""
    try:
        responsible_id = request.args.get('responsible_id')
        week = request.args.get('week', get_current_iso_week())

        print(f"📋 Form accessed - Responsible: {responsible_id}, Week: {week}")

        data = get_responsible_with_kpis(responsible_id, week)
        responsible = data['responsible']
        kpis = data['kpis']

        print(f"📋 Data loaded - Responsible: {responsible['name']}, KPIs: {len(kpis)}")

        if not kpis:
            return '''
            <div style="font-family: Arial; padding: 40px; text-align: center;">
                <h2>ℹ️ No KPIs Found</h2>
                <p>There are no KPIs assigned for this week.</p>
            </div>
            '''

        kpi_html = ""
        for kpi in kpis:
            kpi_html += f'''
            <div class="kpi-card">
                <div class="form-group">
                    <label class="form-label">KPI Name:</label>
                    <input type="text" value="{kpi['KPI_name']}" class="kpi-input" readonly style="background-color:#f8f9fa;" />
                </div>

                <div class="form-group">
                    <label class="form-label">KPI Objective:</label>
                    <input type="text" value="{kpi['KPI_objectif'] or ''}" placeholder="No objective set" class="kpi-input" readonly style="background-color:#f8f9fa;" />
                </div>

                <div class="form-group">
                    <label class="form-label">KPI Value:</label>
                    <input type="text" name="value_{kpi['kpi_values_id']}" value="{kpi['value'] or ''}" placeholder="Enter value" class="kpi-input" readonly style="background-color:#f8f9fa;" />
                </div>

                <div class="form-group">
                    <label class="form-label">Analysis:</label>
                    <textarea name="analyse_{kpi['kpi_values_id']}" placeholder="Enter your analysis here..." class="kpi-textarea">{kpi['analyse'] or ''}</textarea>
                </div>

                <div class="form-group">
                    <label class="form-label">Corrective Actions:</label>
                    <textarea name="actions_{kpi['kpi_values_id']}" placeholder="Enter corrective actions..." class="kpi-textarea">{kpi['actions_correctives'] or ''}</textarea>
                </div>
            </div>
            '''

        return f'''
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <title>KPI Form - Week {week}</title>
            <style>
                body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f4f6f9; padding: 20px; margin: 0; }}
                .container {{ max-width: 900px; margin: 0 auto; background: #fff; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); overflow: hidden; }}
                .header {{ background: #0078D7; color: white; padding: 20px; text-align: center; }}
                .header h1 {{ margin: 0; font-size: 24px; font-weight: 600; }}
                .form-section {{ padding: 30px; }}
                .info-section {{ background: #f8f9fa; padding: 20px; border-radius: 6px; margin-bottom: 25px; border-left: 4px solid #0078D7; }}
                .info-row {{ display: flex; margin-bottom: 15px; align-items: center; }}
                .info-label {{ font-weight: 600; color: #333; width: 120px; font-size: 14px; }}
                .info-value {{ flex: 1; padding: 8px 12px; background: white; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; }}
                .kpi-section {{ margin-top: 30px; }}
                .kpi-section h3 {{ color: #0078D7; margin-bottom: 20px; font-size: 18px; border-bottom: 2px solid #0078D7; padding-bottom: 8px; }}
                .kpi-card {{ background: #fff; border: 1px solid #e1e5e9; border-radius: 6px; padding: 20px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }}
                .form-group {{ margin-bottom: 15px; }}
                .form-label {{ display: block; font-weight: 600; color: #333; margin-bottom: 5px; font-size: 14px; }}
                .kpi-input {{ width: 100%; padding: 10px 12px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; transition: border-color 0.2s; }}
                .kpi-textarea {{ width: 100%; padding: 10px 12px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; min-height: 80px; resize: vertical; transition: border-color 0.2s; font-family: inherit; }}
                .kpi-input:focus, .kpi-textarea:focus {{ border-color: #0078D7; outline: none; box-shadow: 0 0 0 2px rgba(0,120,215,0.1); }}
                .submit-btn {{ background: #0078D7; color: white; border: none; padding: 12px 30px; border-radius: 4px; font-size: 16px; font-weight: 600; cursor: pointer; transition: background-color 0.2s; display: block; width: 100%; margin-top: 20px; }}
                .submit-btn:hover {{ background: #005ea6; }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h1>📊 KPI Submission Form - Week {week}</h1>
                </div>

                <div class="form-section">
                    <div class="info-section">
                        <div class="info-row">
                            <div class="info-label">Responsible</div>
                            <div class="info-value">{responsible['name']}</div>
                        </div>
                        <div class="info-row">
                            <div class="info-label">Week</div>
                            <div class="info-value">{week}</div>
                        </div>
                    </div>

                    <div class="kpi-section">
                        <h3>Your KPIs</h3>
                        <form action="/submit" method="POST">
                            <input type="hidden" name="responsible_id" value="{responsible_id}" />
                            <input type="hidden" name="week" value="{week}" />
                            {kpi_html}
                            <button type="submit" class="submit-btn">Submit KPI Report</button>
                        </form>
                    </div>
                </div>
            </div>
        </body>
        </html>
        '''
    except Exception as e:
        print(f"❌ Error in form_page: {str(e)}")
        traceback.print_exc()
        return f'<p style="color:red; padding: 20px;">Error loading form: {str(e)}</p>'

@app.route('/submit', methods=['POST'])
def submit_form():
    """Handle form submission"""
    try:
        responsible_id = request.form.get('responsible_id')
        week = request.form.get('week')

        print(f"📝 Form submission - Responsible: {responsible_id}, Week: {week}")

        # Collect analyse_* and actions_* fields
        kpi_data = {}
        for key, value in request.form.items():
            if key.startswith('analyse_'):
                kpi_values_id = key.split('_', 1)[1]
                kpi_data.setdefault(kpi_values_id, {})['analyse'] = value
            elif key.startswith('actions_'):
                kpi_values_id = key.split('_', 1)[1]
                kpi_data.setdefault(kpi_values_id, {})['actions_correctives'] = value

        if not kpi_data:
            return '''
            <div style="font-family: Arial; padding: 40px; text-align: center;">
                <h2 style="color:#e67e22;">ℹ️ Nothing to Update</h2>
                <p>No analysis or corrective actions were provided.</p>
            </div>
            ''', 200

        conn = db_pool.getconn()
        try:
            with conn.cursor() as cur:
                for kpi_values_id, data in kpi_data.items():
                    cur.execute(
                        'SELECT "analyse", actions_correctives FROM public.kpi_values WHERE kpi_values_id = %s',
                        (kpi_values_id,),
                    )
                    old = cur.fetchone()
                    if not old:
                        continue

                    old_analyse, old_actions = old
                    new_analyse = data.get('analyse', old_analyse)
                    new_actions = data.get('actions_correctives', old_actions)

                    cur.execute(
                        """
                        UPDATE public.kpi_values
                        SET "analyse" = %s, actions_correctives = %s
                        WHERE kpi_values_id = %s
                        """,
                        (new_analyse, new_actions, kpi_values_id),
                    )

            conn.commit()
            print(f"✅ Successfully updated {len(kpi_data)} KPI value(s)")

            return f'''
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <title>Submission Successful</title>
                <style>
                    body {{
                        font-family:'Segoe UI',sans-serif; background:#f4f4f4;
                        display:flex; justify-content:center; align-items:center;
                        height:100vh; margin:0;
                    }}
                    .success-container {{
                        background:#fff; padding:40px; border-radius:10px;
                        box-shadow:0 4px 15px rgba(0,0,0,0.1); text-align:center;
                        max-width: 500px;
                    }}
                    h1 {{ color:#28a745; font-size:28px; margin-bottom:20px; }}
                    p {{ font-size:16px; color:#333; margin-bottom:30px; }}
                    a {{ display:inline-block; padding:12px 25px; background:#0078D7;
                        color:white; text-decoration:none; border-radius:6px; font-weight:bold; }}
                    a:hover {{ background:#005ea6; }}
                </style>
            </head>
            <body>
                <div class="success-container">
                    <h1>✅ KPI Report Submitted!</h1>
                    <p>Your KPI analysis and corrective actions for week <strong>{week}</strong> have been successfully saved.</p>
                    <p>Thank you for your submission!</p>
                    <a href="/form?responsible_id={responsible_id}&week={week}">View Form Again</a>
                </div>
            </body>
            </html>
            '''
        finally:
            db_pool.putconn(conn)

    except Exception as e:
        print(f"❌ Error in submit_form: {str(e)}")
        traceback.print_exc()
        return f'<h2 style="color:red; padding: 20px;">❌ Failed to submit KPI values</h2><p>{str(e)}</p>', 500

# ========================================
# INITIALIZE SCHEDULER
# ========================================

scheduler = BackgroundScheduler(timezone=pytz.timezone('Africa/Tunis'))
scheduler.add_job(
    scheduled_email_task,
    'cron',
    hour=23,
    minute=35,
    timezone=pytz.timezone('Africa/Tunis'),
    id='kpi_email_scheduler',
    name='KPI Automated Email Scheduler'
)
scheduler.start()

print("\n" + "="*70)
print("✅ KPI AUTOMATION SYSTEM INITIALIZED")
print("="*70)
print(f"📅 Scheduler: Active")
print(f"⏰ Schedule: Daily at 10:15 AM (Africa/Tunis)")
print(f"📧 Email Server: {SMTP_SERVER}:{SMTP_PORT} (Authenticated SMTP)")
print(f"📧 Next run: {scheduler.get_jobs()[0].next_run_time}")
print(f"🌐 Server: Running on port {PORT}")
print("="*70 + "\n")

# ========================================
# START SERVER
# ========================================

if __name__ == '__main__':
    try:
        print(f"🔗 Access points:")
        print(f"   - Home: http://localhost:{PORT}/")
        print(f"   - Form: http://localhost:{PORT}/form?responsible_id=1&week=2025-W43")
        print("\n" + "="*70 + "\n")
        app.run(host='0.0.0.0', port=PORT, debug=False)
    except KeyboardInterrupt:
        print("\n🛑 Server stopped by user")
    finally:
        db_pool.closeall()
        scheduler.shutdown()
        print("✅ Cleanup complete")
