import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import psycopg2
from psycopg2 import pool
from flask import Flask, request, redirect
from apscheduler.schedulers.background import BackgroundScheduler
import pytz
import traceback
from urllib.parse import quote_plus

app = Flask(__name__)

# ---------- Configuration ----------
PORT = int(os.getenv('PORT', 5000))

# Power BI Dashboard URL
POWER_BI_URL = "https://app.powerbi.com/groups/me/reports/9728953d-29c3-4009-99c6-3f61940eb937/b5462e0103e0e04ee51b?ctid=4e99b5ff-dd77-418a-8b69-1d684e911168&experience=power-bi"

# ---------- PostgreSQL Connection Pool ----------
db_pool = pool.SimpleConnectionPool(
    1, 20,
    user="administrationSTS",
    host="avo-adb-002.postgres.database.azure.com",
    database="Subsidy_DB",
    password="St$@0987",
    port=5432,
    sslmode="require"
)

# ---------- Email Configuration ----------
SMTP_SERVER = "avocarbon-com.mail.protection.outlook.com"
SMTP_PORT = 25
EMAIL_USER = "administration.STS@avocarbon.com"
EMAIL_PASSWORD = "shnlgdyfbcztbhxn"

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

def get_responsible_with_kpis(responsible_id, week, plant_id=None):
    """Fetch responsible info and their KPIs for a given week and optionally a specific plant"""
    conn = db_pool.getconn()
    try:
        with conn.cursor() as cur:
            # Fetch responsible info with plant name
            cur.execute(
                """
                SELECT r.responsible_id, r.name, r.email, p.name as plant_name, p.plant_id
                FROM public."Responsible" r
                LEFT JOIN public.plants p ON r.plant_id = p.plant_id
                WHERE r.responsible_id = %s
                """,
                (responsible_id,),
            )
            responsible = cur.fetchone()
            if not responsible:
                raise Exception("Responsible not found")

            # Fetch KPIs - filter by plant_id if provided
            if plant_id:
                cur.execute(
                    """
                    SELECT kv.kpi_values_id, kv.value, kv.week, kv.analyse, kv.actions_correctives,
                           k.kpi_id, k."KPI_name", k."KPI_objectif"
                    FROM public.kpi_values kv
                    JOIN public."Kpi" k ON kv.kpi_id = k.kpi_id
                    WHERE kv.responsible_id = %s AND kv.week = %s AND kv.plant_id = %s
                    ORDER BY k.kpi_id ASC
                    """,
                    (responsible_id, week, plant_id),
                )
            else:
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
                    'plant_name': responsible[3] or 'N/A',
                    'plant_id': responsible[4],
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

def get_all_kpi_values():
    """Fetch all KPI values with responsible, plant, and KPI details"""
    conn = db_pool.getconn()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT 
                    kv.kpi_values_id,
                    r.responsible_id,
                    r.name as responsible_name,
                    p.name as plant_name,
                    k.kpi_id,
                    k."KPI_name",
                    kv.value,
                    kv.week,
                    kv.analyse,
                    kv.actions_correctives,
                    k."KPI_objectif"
                FROM public.kpi_values kv
                JOIN public."Responsible" r ON kv.responsible_id = r.responsible_id
                LEFT JOIN public.plants p ON r.plant_id = p.plant_id
                JOIN public."Kpi" k ON kv.kpi_id = k.kpi_id
                ORDER BY kv.week DESC, r.name ASC, k."KPI_name" ASC
                """
            )
            results = cur.fetchall()
            
            return [
                {
                    'kpi_values_id': row[0],
                    'responsible_id': row[1],
                    'responsible_name': row[2],
                    'plant_name': row[3] or 'N/A',
                    'kpi_id': row[4],
                    'kpi_name': row[5],
                    'value': row[6],
                    'week': row[7],
                    'analyse': row[8],
                    'actions_correctives': row[9],
                    'kpi_objectif': row[10]
                }
                for row in results
            ]
    except Exception as e:
        print(f"❌ Database error in get_all_kpi_values: {str(e)}")
        traceback.print_exc()
        return []
    finally:
        db_pool.putconn(conn)

def send_kpi_email(responsible_id, responsible_name, responsible_email, kpi_name, week, plant_name, plant_id):
    """Send KPI email with a link to the form for a specific responsible and plant"""
    try:
        base = _base_url()
        form_link = f"{base}/form?responsible_id={responsible_id}&week={quote_plus(week)}&plant_id={plant_id}"

        msg = MIMEMultipart()
        msg['From'] = f'"Administration STS" <{EMAIL_USER}>'
        msg['To'] = responsible_email
        msg['Subject'] = f"KPI Report - {kpi_name} - {plant_name} - Week {week}"

        html_content = f"""
        <!DOCTYPE html>
        <html>
        <body style="font-family:Arial,sans-serif; background:#f7f7f7; padding:24px;">
          <div style="max-width:640px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;border:1px solid #eee">
            <div style="background:#0078D7;color:#fff;padding:20px 24px">
              <h2 style="margin:0;font-weight:600">KPI Report – {responsible_name}</h2>
              <div style="margin-top:8px;font-size:14px">Week {week} | {plant_name}</div>
            </div>
            <div style="padding:24px">
              <p>Hello {responsible_name},</p>
              <p>The KPI <strong>{kpi_name}</strong> for <strong>{plant_name}</strong> is due for reporting for week <strong>{week}</strong>.</p>
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

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=10) as server:
            server.send_message(msg)

        print(f"✅ Email sent successfully to {responsible_email} for KPI: {kpi_name} at Plant: {plant_name}")
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
    Fetch all KPIs that are due (frequence_de_envoi <= NOW()) 
    along with their assigned responsibles and plants for the current week.
    Returns: List of tuples (kpi_id, kpi_name, responsible_id, resp_name, email, week, plant_name, plant_id)
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
                    kv.week,
                    COALESCE(p.name, 'N/A') as plant_name,
                    kv.plant_id
                FROM public."Kpi" k
                JOIN public.kpi_values kv ON kv.kpi_id = k.kpi_id
                JOIN public."Responsible" r ON r.responsible_id = kv.responsible_id
                LEFT JOIN public.plants p ON kv.plant_id = p.plant_id
                WHERE k.frequence_de_envoi <= NOW()
                  AND kv.week = %s
                ORDER BY kv.plant_id, k.kpi_id, r.responsible_id
                """,
                (current_week,),
            )

            results = cur.fetchall()
            print(f"📊 Found {len(results)} KPI-Responsible-Plant combinations due for week {current_week}")
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
    3. Groups by plant and sends separate email per plant
    4. Updates created_at to trigger next cycle calculation
    """
    print(f"\n{'='*70}")
    print(f"⏰ SCHEDULED TASK RUNNING at {datetime.now()}")
    print(f"{'='*70}")

    current_week = get_current_iso_week()
    print(f"📅 Current ISO Week: {current_week}")

    # Get all due KPIs with their responsibles and plants
    due_records = get_due_kpis_with_responsibles()

    if not due_records:
        print("ℹ️  No KPIs are due for sending at this time.")
        print(f"{'='*70}\n")
        return

    # Group by (responsible_id, plant_id) to send one email per plant
    plant_groups = {}
    for kpi_id, kpi_name, responsible_id, resp_name, email, week, plant_name, plant_id in due_records:
        key = (responsible_id, plant_id)
        if key not in plant_groups:
            plant_groups[key] = {
                'responsible_id': responsible_id,
                'resp_name': resp_name,
                'email': email,
                'week': week,
                'plant_name': plant_name,
                'plant_id': plant_id,
                'kpis': []
            }
        plant_groups[key]['kpis'].append({
            'kpi_id': kpi_id,
            'kpi_name': kpi_name
        })

    kpis_processed = set()
    emails_sent = 0
    emails_failed = 0

    print(f"\n📧 Processing {len(plant_groups)} plant-responsible combination(s):\n")

    # Send one email per plant
    for key, group_data in plant_groups.items():
        responsible_id = group_data['responsible_id']
        resp_name = group_data['resp_name']
        email = group_data['email']
        week = group_data['week']
        plant_name = group_data['plant_name']
        plant_id = group_data['plant_id']
        kpis = group_data['kpis']
        
        # Use the first KPI name for the email subject (or you could list all)
        kpi_name = kpis[0]['kpi_name']
        if len(kpis) > 1:
            kpi_name = f"{kpi_name} and {len(kpis)-1} more"

        print(f"📤 Sending KPI reminder:")
        print(f"   Plant: {plant_name} (ID: {plant_id})")
        print(f"   KPIs: {', '.join([k['kpi_name'] for k in kpis])}")
        print(f"   To: {resp_name} ({email})")
        print(f"   Week: {week}")

        try:
            success = send_kpi_email(responsible_id, resp_name, email, kpi_name, week, plant_name, plant_id)

            if success:
                emails_sent += 1
                # Mark all KPIs in this email as processed
                for kpi in kpis:
                    kpis_processed.add(kpi['kpi_id'])
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
            .button {{ display: inline-block; margin-top: 20px; padding: 12px 24px; background: #0078D7; color: white; text-decoration: none; border-radius: 6px; font-weight: 600; transition: background 0.2s; }}
            .button:hover {{ background: #005ea6; }}
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
            </div>
            <p>The system automatically checks for due KPIs and sends email notifications to responsible parties <strong>grouped by plant</strong> based on the <code>frequence_de_envoi</code> schedule.</p>
            <a href="/dashboard" class="button">📊 View Dashboard</a>
        </div>
    </body>
    </html>
    '''

@app.route('/dashboard')
def dashboard():
    """Redirect to Power BI dashboard"""
    return redirect(POWER_BI_URL)

@app.route('/scheduler-status')
def scheduler_status():
    """Check scheduler status and jobs"""
    jobs = scheduler.get_jobs()
    jobs_info = []
    
    for job in jobs:
        jobs_info.append({
            'id': job.id,
            'name': job.name,
            'next_run': str(job.next_run_time),
            'trigger': str(job.trigger)
        })
    
    current_time = datetime.now(pytz.timezone('Africa/Tunis'))
    
    return f'''
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>Scheduler Status</title>
        <style>
            body {{ font-family: 'Segoe UI', sans-serif; background: #f4f6f9; padding: 40px; }}
            .container {{ max-width: 900px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
            h1 {{ color: #0078D7; }}
            .info {{ background: #e7f3ff; padding: 15px; border-radius: 5px; margin: 15px 0; border-left: 4px solid #0078D7; }}
            .job {{ background: #f8f9fa; padding: 15px; margin: 10px 0; border-radius: 5px; border-left: 4px solid #28a745; }}
            .label {{ font-weight: 600; color: #333; }}
            pre {{ background: #f4f4f4; padding: 10px; border-radius: 4px; overflow-x: auto; }}
            .button {{ display: inline-block; margin: 10px 5px; padding: 10px 20px; background: #0078D7; color: white; text-decoration: none; border-radius: 6px; font-weight: 600; }}
            .button:hover {{ background: #005ea6; }}
            .test-btn {{ background: #28a745; }}
            .test-btn:hover {{ background: #218838; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>⏰ Scheduler Status</h1>
            
            <div class="info">
                <div><span class="label">Scheduler Running:</span> {scheduler.running}</div>
                <div><span class="label">Current Server Time (UTC):</span> {datetime.now()}</div>
                <div><span class="label">Current Time (Africa/Tunis):</span> {current_time}</div>
                <div><span class="label">Total Jobs:</span> {len(jobs)}</div>
            </div>
            
            <h2>📋 Scheduled Jobs</h2>
            {''.join([f'''
            <div class="job">
                <div><span class="label">Job ID:</span> {job['id']}</div>
                <div><span class="label">Name:</span> {job['name']}</div>
                <div><span class="label">Next Run:</span> {job['next_run']}</div>
                <div><span class="label">Trigger:</span> {job['trigger']}</div>
            </div>
            ''' for job in jobs_info])}
            
            <h2>🧪 Test Functions</h2>
            <a href="/test-email-task" class="button test-btn">🧪 Run Email Task Manually</a>
            <a href="/test-due-kpis" class="button test-btn">📊 Check Due KPIs</a>
            <a href="/" class="button">← Back to Home</a>
        </div>
    </body>
    </html>
    '''

@app.route('/test-email-task')
def test_email_task():
    """Manually trigger the email task for testing"""
    try:
        print("\n" + "="*70)
        print("🧪 MANUAL TEST TRIGGERED from /test-email-task")
        print("="*70)
        scheduled_email_task()
        return '''
        <div style="font-family: Arial; padding: 40px; text-align: center;">
            <h2 style="color: #28a745;">✅ Email Task Executed</h2>
            <p>Check the server logs for results.</p>
            <a href="/scheduler-status" style="display: inline-block; margin-top: 20px; padding: 12px 24px; background: #0078D7; color: white; text-decoration: none; border-radius: 6px;">Back to Scheduler Status</a>
        </div>
        '''
    except Exception as e:
        return f'''
        <div style="font-family: Arial; padding: 40px; text-align: center;">
            <h2 style="color: #dc3545;">❌ Error</h2>
            <p>{str(e)}</p>
            <pre style="text-align: left; background: #f4f4f4; padding: 15px; border-radius: 5px;">{traceback.format_exc()}</pre>
            <a href="/scheduler-status" style="display: inline-block; margin-top: 20px; padding: 12px 24px; background: #0078D7; color: white; text-decoration: none; border-radius: 6px;">Back to Scheduler Status</a>
        </div>
        '''

@app.route('/test-due-kpis')
def test_due_kpis():
    """Check which KPIs are currently due"""
    try:
        current_week = get_current_iso_week()
        due_records = get_due_kpis_with_responsibles()
        
        records_html = ""
        if due_records:
            for kpi_id, kpi_name, responsible_id, resp_name, email, week, plant_name, plant_id in due_records:
                records_html += f'''
                <div style="background: #f8f9fa; padding: 15px; margin: 10px 0; border-radius: 5px; border-left: 4px solid #0078D7;">
                    <div><strong>KPI:</strong> {kpi_name} (ID: {kpi_id})</div>
                    <div><strong>Responsible:</strong> {resp_name} (ID: {responsible_id})</div>
                    <div><strong>Email:</strong> {email}</div>
                    <div><strong>Plant:</strong> {plant_name} (ID: {plant_id})</div>
                    <div><strong>Week:</strong> {week}</div>
                </div>
                '''
        else:
            records_html = '<p style="color: #666;">No KPIs are currently due.</p>'
        
        return f'''
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <title>Due KPIs Check</title>
            <style>
                body {{ font-family: 'Segoe UI', sans-serif; background: #f4f6f9; padding: 40px; }}
                .container {{ max-width: 900px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
                h1 {{ color: #0078D7; }}
                .info {{ background: #e7f3ff; padding: 15px; border-radius: 5px; margin: 15px 0; border-left: 4px solid #0078D7; }}
                .button {{ display: inline-block; margin-top: 20px; padding: 12px 24px; background: #0078D7; color: white; text-decoration: none; border-radius: 6px; font-weight: 600; }}
                .button:hover {{ background: #005ea6; }}
            </style>
        </head>
        <body>
            <div class="container">
                <h1>📊 Due KPIs Check</h1>
                <div class="info">
                    <div><strong>Current Week:</strong> {current_week}</div>
                    <div><strong>Total Due KPIs:</strong> {len(due_records)}</div>
                    <div><strong>Check Time:</strong> {datetime.now()}</div>
                </div>
                
                <h2>KPI Records Found:</h2>
                {records_html}
                
                <a href="/scheduler-status" class="button">← Back to Scheduler Status</a>
            </div>
        </body>
        </html>
        '''
    except Exception as e:
        return f'''
        <div style="font-family: Arial; padding: 40px; text-align: center;">
            <h2 style="color: #dc3545;">❌ Error</h2>
            <p>{str(e)}</p>
            <pre style="text-align: left; background: #f4f4f4; padding: 15px; border-radius: 5px;">{traceback.format_exc()}</pre>
        </div>
        '''

@app.route('/form')
def form_page():
    """Display KPI form filtered by plant_id if provided"""
    try:
        responsible_id = request.args.get('responsible_id')
        week = request.args.get('week', get_current_iso_week())
        plant_id = request.args.get('plant_id')  # NEW: Get plant_id from URL

        print(f"📋 Form accessed - Responsible: {responsible_id}, Week: {week}, Plant: {plant_id}")

        # Pass plant_id to filter KPIs
        data = get_responsible_with_kpis(responsible_id, week, plant_id)
        responsible = data['responsible']
        kpis = data['kpis']

        # Get actual plant name from URL parameter if filtering
        actual_plant_name = responsible['plant_name']
        if plant_id:
            conn = db_pool.getconn()
            try:
                with conn.cursor() as cur:
                    cur.execute("SELECT name FROM public.plants WHERE plant_id = %s", (plant_id,))
                    result = cur.fetchone()
                    if result:
                        actual_plant_name = result[0]
            finally:
                db_pool.putconn(conn)

        print(f"📋 Data loaded - Responsible: {responsible['name']}, Plant: {actual_plant_name}, KPIs: {len(kpis)}")

        if not kpis:
            return '''
            <div style="font-family: Arial; padding: 40px; text-align: center;">
                <h2>ℹ️ No KPIs Found</h2>
                <p>There are no KPIs assigned for this week and plant.</p>
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
                    <label class="form-label">Analysis: <span style="color:#999;font-weight:normal;font-size:12px;">(Provide detailed analysis)</span></label>
                    <textarea name="analyse_{kpi['kpi_values_id']}" placeholder="Enter your detailed analysis here..." class="kpi-textarea-large">{kpi['analyse'] or ''}</textarea>
                </div>

                <div class="form-group">
                    <label class="form-label">Corrective Actions: <span style="color:#999;font-weight:normal;font-size:12px;">(Provide detailed corrective actions)</span></label>
                    <textarea name="actions_{kpi['kpi_values_id']}" placeholder="Enter detailed corrective actions here..." class="kpi-textarea-large">{kpi['actions_correctives'] or ''}</textarea>
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
                .container {{ max-width: 1000px; margin: 0 auto; background: #fff; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); overflow: hidden; }}
                .header {{ background: #0078D7; color: white; padding: 25px; text-align: center; }}
                .header h1 {{ margin: 0; font-size: 26px; font-weight: 600; }}
                .header .subtitle {{ margin-top: 8px; font-size: 14px; opacity: 0.9; }}
                .form-section {{ padding: 30px; }}
                .info-section {{ background: #f8f9fa; padding: 20px; border-radius: 6px; margin-bottom: 25px; border-left: 4px solid #0078D7; }}
                .info-row {{ display: flex; margin-bottom: 15px; align-items: center; }}
                .info-label {{ font-weight: 600; color: #333; width: 140px; font-size: 14px; }}
                .info-value {{ flex: 1; padding: 10px 12px; background: white; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; }}
                .kpi-section {{ margin-top: 30px; }}
                .kpi-section h3 {{ color: #0078D7; margin-bottom: 20px; font-size: 20px; border-bottom: 2px solid #0078D7; padding-bottom: 10px; }}
                .kpi-card {{ background: #fff; border: 1px solid #e1e5e9; border-radius: 6px; padding: 25px; margin-bottom: 25px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }}
                .form-group {{ margin-bottom: 18px; }}
                .form-label {{ display: block; font-weight: 600; color: #333; margin-bottom: 6px; font-size: 14px; }}
                .kpi-input {{ width: 100%; padding: 10px 12px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; transition: border-color 0.2s; box-sizing: border-box; }}
                .kpi-textarea {{ width: 100%; padding: 10px 12px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; min-height: 100px; resize: vertical; transition: border-color 0.2s; font-family: inherit; box-sizing: border-box; }}
                .kpi-textarea-large {{ width: 100%; padding: 12px 15px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; min-height: 180px; resize: vertical; transition: border-color 0.2s; font-family: inherit; line-height: 1.6; box-sizing: border-box; }}
                .kpi-input:focus, .kpi-textarea:focus, .kpi-textarea-large:focus {{ border-color: #0078D7; outline: none; box-shadow: 0 0 0 2px rgba(0,120,215,0.1); }}
                .submit-btn {{ background: #0078D7; color: white; border: none; padding: 14px 30px; border-radius: 4px; font-size: 16px; font-weight: 600; cursor: pointer; transition: background-color 0.2s; display: block; width: 100%; margin-top: 20px; }}
                .submit-btn:hover {{ background: #005ea6; }}
                .dashboard-btn {{ background: #28a745; color: white; border: none; padding: 14px 30px; border-radius: 4px; font-size: 16px; font-weight: 600; cursor: pointer; transition: background-color 0.2s; display: block; width: 100%; margin-top: 10px; text-decoration: none; text-align: center; }}
                .dashboard-btn:hover {{ background: #218838; }}
                .char-counter {{ font-size: 12px; color: #666; margin-top: 4px; text-align: right; }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h1>📊 KPI Submission Form</h1>
                </div>

                <div class="form-section">
                    <div class="info-section">
                        <div class="info-row">
                            <div class="info-label">Responsible:</div>
                            <div class="info-value">{responsible['name']}</div>
                        </div>
                        <div class="info-row">
                            <div class="info-label">Plant:</div>
                            <div class="info-value">{actual_plant_name}</div>
                        </div>
                        <div class="info-row">
                            <div class="info-label">Week:</div>
                            <div class="info-value">{week}</div>
                        </div>
                    </div>

                    <div class="kpi-section">
                        <h3>Your KPIs for {actual_plant_name}</h3>
                        <form action="/submit" method="POST">
                            <input type="hidden" name="responsible_id" value="{responsible_id}" />
                            <input type="hidden" name="week" value="{week}" />
                            <input type="hidden" name="plant_id" value="{plant_id or ''}" />
                            {kpi_html}
                            <button type="submit" class="submit-btn">📤 Submit KPI Report</button>
                        </form>
                    </div>
                </div>
            </div>
            
            <script>
                // Add character counters for textareas
                document.addEventListener('DOMContentLoaded', function() {{
                    const textareas = document.querySelectorAll('.kpi-textarea-large');
                    textareas.forEach(textarea => {{
                        const counter = document.createElement('div');
                        counter.className = 'char-counter';
                        counter.textContent = textarea.value.length + ' characters';
                        textarea.parentNode.appendChild(counter);
                        
                        textarea.addEventListener('input', function() {{
                            counter.textContent = this.value.length + ' characters';
                        }});
                    }});
                }});
            </script>
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
        plant_id = request.form.get('plant_id')  # NEW: Get plant_id from form

        print(f"📝 Form submission - Responsible: {responsible_id}, Week: {week}, Plant: {plant_id}")

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

            # Build redirect URL with plant_id if provided
            form_url = f"/form?responsible_id={responsible_id}&week={week}"
            if plant_id:
                form_url += f"&plant_id={plant_id}"

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
                        background:#fff; padding:50px; border-radius:10px;
                        box-shadow:0 4px 15px rgba(0,0,0,0.1); text-align:center;
                        max-width: 550px;
                    }}
                    .success-icon {{ font-size: 64px; margin-bottom: 20px; }}
                    h1 {{ color:#28a745; font-size:28px; margin-bottom:20px; }}
                    p {{ font-size:16px; color:#333; margin-bottom:30px; line-height: 1.6; }}
                    .button-group {{ display: flex; gap: 10px; justify-content: center; flex-wrap: wrap; }}
                    a {{ display:inline-block; padding:12px 25px; background:#0078D7;
                        color:white; text-decoration:none; border-radius:6px; font-weight:600; 
                        margin: 5px; transition: all 0.2s; }}
                    a:hover {{ background:#005ea6; transform: translateY(-2px); }}
                    a.dashboard {{ background:#28a745; }}
                    a.dashboard:hover {{ background:#218838; }}
                </style>
            </head>
            <body>
                <div class="success-container">
                    <div class="success-icon">✅</div>
                    <h1>KPI Report Submitted Successfully!</h1>
                    <p>Your KPI analysis and corrective actions for week <strong>{week}</strong> have been successfully saved to the system.</p>
                    <p>Thank you for your timely submission!</p>
                    <div class="button-group">
                        <a href="{form_url}">🔄 View Form Again</a>
                        <a href="/dashboard" class="dashboard">📊 View Dashboard</a>
                    </div>
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
    hour=7,
    minute=14,
    timezone=pytz.timezone('Africa/Tunis'),
    id='kpi_email_scheduler',
    name='KPI Automated Email Scheduler (Plant-based)'
)
scheduler.start()

print("\n" + "="*70)
print("✅ KPI AUTOMATION SYSTEM INITIALIZED (PLANT-BASED EMAILS)")
print("="*70)
print(f"📅 Scheduler: Active")
print(f"⏰ Schedule: Daily at 07:05 AM (Africa/Tunis)")
print(f"📧 Next run: {scheduler.get_jobs()[0].next_run_time}")
print(f"🌐 Server: Running on port {PORT}")
print(f"🏭 Mode: One email per plant (grouped KPIs)")
print("="*70 + "\n")

# ========================================
# START SERVER
# ========================================

if __name__ == '__main__':
    try:
        print(f"🔗 Access points:")
        print(f"   - Home: http://localhost:{PORT}/")
        print(f"   - Dashboard: http://localhost:{PORT}/dashboard (redirects to Power BI)")
        print(f"   - Form: http://localhost:{PORT}/form?responsible_id=1&week=2025-W45&plant_id=1")
        print(f"   - Scheduler Status: http://localhost:{PORT}/scheduler-status")
        print("\n" + "="*70 + "\n")
        app.run(host='0.0.0.0', port=PORT, debug=False)
    except KeyboardInterrupt:
        print("\n🛑 Server stopped by user")
    finally:
        db_pool.closeall()
        scheduler.shutdown()
        print("✅ Cleanup complete")
