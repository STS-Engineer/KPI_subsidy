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
import logging

# Configure logging for Azure
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# ---------- Configuration ----------
PORT = int(os.getenv('PORT', 8000))
APP_BASE_URL = os.getenv('WEBSITE_HOSTNAME', 'kpi-subsidy.azurewebsites.net')

# ---------- PostgreSQL Connection Pool ----------
try:
    db_pool = pool.SimpleConnectionPool(
        1, 20,
        user="administrationSTS",
        host="avo-adb-002.postgres.database.azure.com",
        database="Subsidy_DB",
        password="St$@0987",
        port=5432,
        sslmode="require",
        connect_timeout=10
    )
    logger.info("✅ Database pool initialized successfully")
except Exception as e:
    logger.error(f"❌ Failed to initialize database pool: {str(e)}")
    db_pool = None

# ---------- Email Configuration ----------
SMTP_SERVER = "avocarbon-com.mail.protection.outlook.com"
SMTP_PORT = 25
EMAIL_USER = "administration.STS@avocarbon.com"
EMAIL_PASSWORD = "shnlgdyfbcztbhxn"

# ========================================
# HELPER FUNCTIONS
# ========================================

def _base_url():
    """Get base URL for email links - Azure compatible"""
    if APP_BASE_URL:
        return f"https://{APP_BASE_URL}"
    return f"http://localhost:{PORT}"

def get_current_iso_week():
    """Get current ISO week in format YYYY-Wxx (e.g., 2025-W43)"""
    now = datetime.now(pytz.timezone('Africa/Tunis'))
    iso_calendar = now.isocalendar()
    return f"{iso_calendar[0]}-W{iso_calendar[1]:02d}"

def get_db_connection():
    """Get database connection with error handling"""
    if not db_pool:
        raise Exception("Database pool not initialized")
    return db_pool.getconn()

def return_db_connection(conn):
    """Return database connection to pool"""
    if db_pool and conn:
        db_pool.putconn(conn)

def get_responsible_with_kpis(responsible_id, week):
    """Fetch responsible info and their KPIs for a given week"""
    conn = None
    try:
        conn = get_db_connection()
        with conn.cursor() as cur:
            # Fetch responsible info with plant name
            cur.execute(
                """
                SELECT r.responsible_id, r.name, r.email, p.name as plant_name
                FROM public."Responsible" r
                LEFT JOIN public.plants p ON r.plant_id = p.plant_id
                WHERE r.responsible_id = %s
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
                    'plant_name': responsible[3] or 'N/A',
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
        logger.error(f"❌ Database error in get_responsible_with_kpis: {str(e)}")
        raise
    finally:
        return_db_connection(conn)

def get_all_kpi_values():
    """Fetch all KPI values with responsible, plant, and KPI details"""
    conn = None
    try:
        conn = get_db_connection()
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
        logger.error(f"❌ Database error in get_all_kpi_values: {str(e)}")
        traceback.print_exc()
        return []
    finally:
        return_db_connection(conn)

def send_kpi_email(responsible_id, responsible_name, responsible_email, kpi_name, week, plant_name):
    """Send KPI email with a link to the form for a specific responsible"""
    try:
        base = _base_url()
        form_link = f"{base}/form?responsible_id={responsible_id}&week={quote_plus(week)}"

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

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as server:
            server.send_message(msg)

        logger.info(f"✅ Email sent successfully to {responsible_email} for KPI: {kpi_name}")
        return True

    except Exception as e:
        logger.error(f"❌ Failed to send email to {responsible_email}: {str(e)}")
        traceback.print_exc()
        return False

# ========================================
# SCHEDULER FUNCTIONS
# ========================================

def get_due_kpis_with_responsibles():
    """Fetch all KPIs that are due for the current week"""
    conn = None
    current_week = get_current_iso_week()

    try:
        conn = get_db_connection()
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
                    COALESCE(p.name, 'N/A') as plant_name
                FROM public."Kpi" k
                JOIN public.kpi_values kv ON kv.kpi_id = k.kpi_id
                JOIN public."Responsible" r ON r.responsible_id = kv.responsible_id
                LEFT JOIN public.plants p ON r.plant_id = p.plant_id
                WHERE k.frequence_de_envoi <= NOW()
                  AND kv.week = %s
                ORDER BY k.kpi_id, r.responsible_id
                """,
                (current_week,),
            )

            results = cur.fetchall()
            logger.info(f"📊 Found {len(results)} KPI-Responsible combinations due for week {current_week}")
            return results

    except Exception as e:
        logger.error(f"❌ Error fetching due KPIs: {str(e)}")
        traceback.print_exc()
        return []
    finally:
        return_db_connection(conn)

def update_kpi_created_at(kpi_id):
    """Update created_at = NOW() to trigger recalculation of frequence_de_envoi"""
    conn = None
    try:
        conn = get_db_connection()
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
                logger.info(f"   ✅ Updated '{kpi_name}' (ID:{kpi_id})")
                logger.info(f"      New created_at: {new_created}")
                logger.info(f"      Next send scheduled: {new_freq}")
                return True
            else:
                logger.warning(f"   ⚠️ KPI {kpi_id} not found for update")
                return False

    except Exception as e:
        logger.error(f"   ❌ Error updating KPI {kpi_id}: {str(e)}")
        traceback.print_exc()
        if conn:
            conn.rollback()
        return False
    finally:
        return_db_connection(conn)

def scheduled_email_task():
    """Main automated scheduler task"""
    logger.info("\n" + "="*70)
    logger.info(f"⏰ SCHEDULED TASK RUNNING at {datetime.now(pytz.timezone('Africa/Tunis'))}")
    logger.info("="*70)

    current_week = get_current_iso_week()
    logger.info(f"📅 Current ISO Week: {current_week}")

    # Get all due KPIs with their responsibles
    due_records = get_due_kpis_with_responsibles()

    if not due_records:
        logger.info("ℹ️  No KPIs are due for sending at this time.")
        logger.info("="*70 + "\n")
        return

    kpis_processed = set()
    emails_sent = 0
    emails_failed = 0

    logger.info(f"\n📧 Processing {len(due_records)} KPI-Responsible combination(s):\n")

    # Send email to each responsible for their due KPIs
    for kpi_id, kpi_name, responsible_id, resp_name, email, week, plant_name in due_records:
        logger.info(f"📤 Sending KPI reminder:")
        logger.info(f"   KPI: {kpi_name} (ID: {kpi_id})")
        logger.info(f"   To: {resp_name} ({email})")
        logger.info(f"   Plant: {plant_name}")
        logger.info(f"   Week: {week}")

        try:
            success = send_kpi_email(responsible_id, resp_name, email, kpi_name, week, plant_name)

            if success:
                emails_sent += 1
                kpis_processed.add(kpi_id)
                logger.info(f"   ✅ Email sent successfully\n")
            else:
                emails_failed += 1
                logger.info(f"   ❌ Email failed to send\n")

        except Exception as e:
            emails_failed += 1
            logger.error(f"   ❌ Exception: {str(e)}\n")
            traceback.print_exc()

    # Update KPIs to schedule next send
    if kpis_processed:
        logger.info(f"\n🔄 Updating {len(kpis_processed)} KPI(s) for next cycle:\n")

        for kpi_id in kpis_processed:
            update_kpi_created_at(kpi_id)

    logger.info(f"\n{'='*70}")
    logger.info(f"✅ TASK COMPLETED:")
    logger.info(f"   📧 Emails sent: {emails_sent}")
    logger.info(f"   ❌ Emails failed: {emails_failed}")
    logger.info(f"   🔄 KPIs updated: {len(kpis_processed)}")
    logger.info("="*70 + "\n")

# ========================================
# FLASK ROUTES
# ========================================

@app.route('/')
def home():
    """Home page with system status"""
    try:
        next_run = scheduler.get_jobs()[0].next_run_time if scheduler.get_jobs() else "Not scheduled"
    except:
        next_run = "Scheduler initializing..."
    
    return f'''
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>KPI Automation System</title>
        <style>
            body {{ font-family: 'Segoe UI', sans-serif; background: #f4f6f9; padding: 20px; margin: 0; }}
            .container {{ max-width: 800px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
            h1 {{ color: #0078D7; margin-bottom: 20px; }}
            .status {{ background: #e7f3ff; padding: 15px; border-radius: 5px; margin: 20px 0; border-left: 4px solid #0078D7; }}
            .info {{ margin: 10px 0; }}
            .label {{ font-weight: 600; color: #333; }}
            .button {{ display: inline-block; margin-top: 20px; padding: 12px 24px; background: #0078D7; color: white; text-decoration: none; border-radius: 6px; font-weight: 600; transition: background 0.2s; }}
            .button:hover {{ background: #005ea6; }}
            .footer {{ margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee; font-size: 12px; color: #666; text-align: center; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>🎯 KPI Automation System</h1>
            <div class="status">
                <div class="info"><span class="label">Status:</span> ✅ Running on Azure</div>
                <div class="info"><span class="label">Current Week:</span> {get_current_iso_week()}</div>
                <div class="info"><span class="label">Next Scheduled Check:</span> {next_run}</div>
                <div class="info"><span class="label">Server Time (Africa/Tunis):</span> {datetime.now(pytz.timezone('Africa/Tunis')).strftime('%Y-%m-%d %H:%M:%S')}</div>
                <div class="info"><span class="label">Base URL:</span> {_base_url()}</div>
            </div>
            <p>The system automatically checks for due KPIs and sends email notifications to responsible parties based on the <code>frequence_de_envoi</code> schedule.</p>
            <a href="/dashboard" class="button">📊 View Dashboard</a>
            <div class="footer">
                Deployed on Azure Web App | Version 2.0
            </div>
        </div>
    </body>
    </html>
    '''

@app.route('/health')
def health_check():
    """Health check endpoint for Azure"""
    try:
        # Test database connection
        conn = get_db_connection()
        return_db_connection(conn)
        
        return {
            'status': 'healthy',
            'timestamp': datetime.now(pytz.timezone('Africa/Tunis')).isoformat(),
            'database': 'connected',
            'scheduler': 'active' if scheduler.running else 'inactive'
        }, 200
    except Exception as e:
        logger.error(f"Health check failed: {str(e)}")
        return {
            'status': 'unhealthy',
            'error': str(e)
        }, 500

@app.route('/dashboard')
def dashboard():
    """Display KPI values dashboard"""
    try:
        kpi_data = get_all_kpi_values()
        
        if not kpi_data:
            return '''
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>KPI Dashboard</title>
                <style>
                    body { font-family: 'Segoe UI', sans-serif; background: #f4f6f9; padding: 20px; margin: 0; }
                    .container { max-width: 1200px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
                    h1 { color: #0078D7; }
                </style>
            </head>
            <body>
                <div class="container">
                    <h1>📊 KPI Dashboard</h1>
                    <p>No KPI data available.</p>
                    <a href="/" style="color: #0078D7;">← Back to Home</a>
                </div>
            </body>
            </html>
            '''
        
        # Generate table rows
        rows_html = ""
        for item in kpi_data:
            analyse_text = (item['analyse'] or 'N/A').replace('<', '&lt;').replace('>', '&gt;').replace('`', '&#96;')
            actions_text = (item['actions_correctives'] or 'N/A').replace('<', '&lt;').replace('>', '&gt;').replace('`', '&#96;')
            
            rows_html += f'''
            <tr>
                <td>{item['kpi_values_id']}</td>
                <td>{item['responsible_name']}</td>
                <td><strong>{item['plant_name']}</strong></td>
                <td>{item['kpi_name']}</td>
                <td>{item['value'] or 'N/A'}</td>
                <td>{item['week']}</td>
                <td class="text-cell">
                    <div class="text-preview">{analyse_text[:100]}{'' if len(analyse_text) <= 100 else '...'}</div>
                    <button class="view-btn" onclick='showFullText("Analyse - {item['kpi_name']}", `{analyse_text}`)'>View Full</button>
                </td>
                <td class="text-cell">
                    <div class="text-preview">{actions_text[:100]}{'' if len(actions_text) <= 100 else '...'}</div>
                    <button class="view-btn" onclick='showFullText("Actions Correctives - {item['kpi_name']}", `{actions_text}`)'>View Full</button>
                </td>
            </tr>
            '''
        
        return f'''
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>KPI Dashboard</title>
            <style>
                body {{
                    font-family: 'Segoe UI', sans-serif;
                    background: #f4f6f9;
                    padding: 20px;
                    margin: 0;
                }}
                .container {{
                    max-width: 1600px;
                    margin: 0 auto;
                    background: #fff;
                    padding: 30px;
                    border-radius: 10px;
                    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                }}
                .header {{
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    margin-bottom: 30px;
                    padding-bottom: 20px;
                    border-bottom: 2px solid #0078D7;
                    flex-wrap: wrap;
                    gap: 15px;
                }}
                h1 {{
                    color: #0078D7;
                    margin: 0;
                }}
                .back-link {{
                    color: #0078D7;
                    text-decoration: none;
                    font-weight: 600;
                    padding: 10px 20px;
                    border: 2px solid #0078D7;
                    border-radius: 6px;
                    transition: all 0.2s;
                }}
                .back-link:hover {{
                    background: #0078D7;
                    color: white;
                }}
                .stats {{
                    background: #e7f3ff;
                    padding: 15px;
                    border-radius: 6px;
                    margin-bottom: 20px;
                    border-left: 4px solid #0078D7;
                }}
                .table-wrapper {{
                    overflow-x: auto;
                    border: 1px solid #ddd;
                    border-radius: 6px;
                }}
                table {{
                    width: 100%;
                    border-collapse: collapse;
                    min-width: 1400px;
                }}
                thead {{
                    background: #0078D7;
                    color: white;
                }}
                th {{
                    padding: 12px;
                    text-align: left;
                    font-weight: 600;
                    position: sticky;
                    top: 0;
                    background: #0078D7;
                }}
                td {{
                    padding: 12px;
                    border-bottom: 1px solid #eee;
                    vertical-align: top;
                }}
                tr:hover {{
                    background: #f8f9fa;
                }}
                .text-cell {{
                    max-width: 300px;
                }}
                .text-preview {{
                    margin-bottom: 8px;
                    line-height: 1.4;
                    color: #333;
                }}
                .view-btn {{
                    background: #0078D7;
                    color: white;
                    border: none;
                    padding: 6px 12px;
                    border-radius: 4px;
                    cursor: pointer;
                    font-size: 12px;
                    transition: background 0.2s;
                }}
                .view-btn:hover {{
                    background: #005ea6;
                }}
                
                /* Modal styles */
                .modal {{
                    display: none;
                    position: fixed;
                    z-index: 1000;
                    left: 0;
                    top: 0;
                    width: 100%;
                    height: 100%;
                    background-color: rgba(0,0,0,0.5);
                }}
                .modal-content {{
                    background-color: #fff;
                    margin: 5% auto;
                    padding: 30px;
                    border-radius: 10px;
                    width: 90%;
                    max-width: 800px;
                    max-height: 70vh;
                    overflow-y: auto;
                    box-shadow: 0 4px 20px rgba(0,0,0,0.3);
                }}
                .modal-header {{
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    margin-bottom: 20px;
                    padding-bottom: 15px;
                    border-bottom: 2px solid #0078D7;
                }}
                .modal-header h2 {{
                    color: #0078D7;
                    margin: 0;
                }}
                .close {{
                    color: #aaa;
                    font-size: 32px;
                    font-weight: bold;
                    cursor: pointer;
                    line-height: 1;
                }}
                .close:hover {{
                    color: #000;
                }}
                .modal-body {{
                    white-space: pre-wrap;
                    line-height: 1.6;
                    color: #333;
                }}
                
                @media (max-width: 768px) {{
                    .container {{
                        padding: 15px;
                    }}
                    .header {{
                        flex-direction: column;
                    }}
                }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h1>📊 KPI Values Dashboard</h1>
                    <a href="/" class="back-link">← Back to Home</a>
                </div>
                
                <div class="stats">
                    <strong>Total Records:</strong> {len(kpi_data)}
                </div>
                
                <div class="table-wrapper">
                    <table>
                        <thead>
                            <tr>
                                <th>ID</th>
                                <th>Responsible</th>
                                <th>Plant</th>
                                <th>KPI Name</th>
                                <th>Value</th>
                                <th>Week</th>
                                <th>Analysis</th>
                                <th>Corrective Actions</th>
                            </tr>
                        </thead>
                        <tbody>
                            {rows_html}
                        </tbody>
                    </table>
                </div>
            </div>
            
            <!-- Modal -->
            <div id="textModal" class="modal">
                <div class="modal-content">
                    <div class="modal-header">
                        <h2 id="modalTitle"></h2>
                        <span class="close" onclick="closeModal()">&times;</span>
                    </div>
                    <div id="modalBody" class="modal-body"></div>
                </div>
            </div>
            
            <script>
                function showFullText(title, text) {{
                    document.getElementById('modalTitle').textContent = title;
                    document.getElementById('modalBody').textContent = text;
                    document.getElementById('textModal').style.display = 'block';
                }}
                
                function closeModal() {{
                    document.getElementById('textModal').style.display = 'none';
                }}
                
                window.onclick = function(event) {{
                    const modal = document.getElementById('textModal');
                    if (event.target == modal) {{
                        modal.style.display = 'none';
                    }}
                }}
                
                // Close modal on escape key
                document.addEventListener('keydown', function(event) {{
                    if (event.key === 'Escape') {{
                        closeModal();
                    }}
                }});
            </script>
        </body>
        </html>
        '''
        
    except Exception as e:
        logger.error(f"❌ Error in dashboard: {str(e)}")
        traceback.print_exc()
        return f'<p style="color:red; padding: 20px;">Error loading dashboard: {str(e)}</p>'

@app.route('/form')
def form_page():
    """Display KPI form"""
    try:
        responsible_id = request.args.get('responsible_id')
        week = request.args.get('week', get_current_iso_week())

        logger.info(f"📋 Form accessed - Responsible: {responsible_id}, Week: {week}")

        data = get_responsible_with_kpis(responsible_id, week)
        responsible = data['responsible']
        kpis = data['kpis']

        logger.info(f"📋 Data loaded - Responsible: {responsible['name']}, Plant: {responsible['plant_name']}, KPIs: {len(kpis)}")

        if not kpis:
            return '''
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>No KPIs Found</title>
                <style>
                    body { font-family: Arial; padding: 40px; text-align: center; background: #f4f6f9; }
                    .message { background: white; padding: 40px; border-radius: 10px; max-width: 500px; margin: 0 auto; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
                    h2 { color: #0078D7; }
                </style>
            </head>
            <body>
                <div class="message">
                    <h2>ℹ️ No KPIs Found</h2>
                    <p>There are no KPIs assigned for this week.</p>
                    <a href="/" style="color: #0078D7; text-decoration: none; font-weight: 600;">← Back to Home</a>
                </div>
            </body>
            </html>
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
                    <textarea name="analyse_{kpi['kpi_values_id']}" placeholder="Enter your detailed analysis here..." class="kpi-textarea-large" required>{kpi['analyse'] or ''}</textarea>
                </div>

                <div class="form-group">
                    <label class="form-label">Corrective Actions: <span style="color:#999;font-weight:normal;font-size:12px;">(Provide detailed corrective actions)</span></label>
                    <textarea name="actions_{kpi['kpi_values_id']}" placeholder="Enter detailed corrective actions here..." class="kpi-textarea-large" required>{kpi['actions_correctives'] or ''}</textarea>
                </div>
            </div>
            '''

        return f'''
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>KPI Form - Week {week}</title>
            <style>
                body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f4f6f9; padding: 20px; margin: 0; }}
                .container {{ max-width: 1000px; margin: 0 auto; background: #fff; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); overflow: hidden; }}
                .header {{ background: #0078D7; color: white; padding: 25px; text-align: center; }}
                .header h1 {{ margin: 0; font-size: 26px; font-weight: 600; }}
                .header .subtitle {{ margin-top: 8px; font-size: 14px; opacity: 0.9; }}
                .form-section {{ padding: 30px; }}
                .info-section {{ background: #f8f9fa; padding: 20px; border-radius: 6px; margin-bottom: 25px; border-left: 4px solid #0078D7; }}
                .info-row {{ display: flex; margin-bottom: 15px; align-items: center; flex-wrap: wrap; }}
                .info-label {{ font-weight: 600; color: #333; width: 140px; font-size: 14px; }}
                .info-value {{ flex: 1; padding: 10px 12px; background: white; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; min-width: 200px; }}
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
                .submit-btn:disabled {{ background: #ccc; cursor: not-allowed; }}
                .dashboard-btn {{ background: #28a745; color: white; border: none; padding: 14px 30px; border-radius: 4px; font-size: 16px; font-weight: 600; cursor: pointer; transition: background-color 0.2s; display: block; width: 100%; margin-top: 10px; text-decoration: none; text-align: center; }}
                .dashboard-btn:hover {{ background: #218838; }}
                .char-counter {{ font-size: 12px; color: #666; margin-top: 4px; text-align: right; }}
                
                @media (max-width: 768px) {{
                    .container {{ margin: 0; border-radius: 0; }}
                    .form-section {{ padding: 15px; }}
                    .info-row {{ flex-direction: column; align-items: stretch; }}
                    .info-label {{ width: 100%; margin-bottom: 5px; }}
                }}
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
                            <div class="info-value">{responsible['plant_name']}</div>
                        </div>
                        <div class="info-row">
                            <div class="info-label">Week:</div>
                            <div class="info-value">{week}</div>
                        </div>
                    </div>

                    <div class="kpi-section">
                        <h3>Your KPIs</h3>
                        <form action="/submit" method="POST" id="kpiForm">
                            <input type="hidden" name="responsible_id" value="{responsible_id}" />
                            <input type="hidden" name="week" value="{week}" />
                            {kpi_html}
                            <button type="submit" class="submit-btn" id="submitBtn">📤 Submit KPI Report</button>
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
                    
                    // Form validation
                    const form = document.getElementById('kpiForm');
                    form.addEventListener('submit', function(e) {{
                        const submitBtn = document.getElementById('submitBtn');
                        submitBtn.disabled = true;
                        submitBtn.textContent = '⏳ Submitting...';
                    }});
                }});
            </script>
        </body>
        </html>
        '''
    except Exception as e:
        logger.error(f"❌ Error in form_page: {str(e)}")
        traceback.print_exc()
        return f'<p style="color:red; padding: 20px;">Error loading form: {str(e)}</p>'

@app.route('/submit', methods=['POST'])
def submit_form():
    """Handle form submission"""
    conn = None
    try:
        responsible_id = request.form.get('responsible_id')
        week = request.form.get('week')

        logger.info(f"📝 Form submission - Responsible: {responsible_id}, Week: {week}")

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
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Nothing to Update</title>
                <style>
                    body { font-family: Arial; padding: 40px; text-align: center; background: #f4f6f9; }
                    .message { background: white; padding: 40px; border-radius: 10px; max-width: 500px; margin: 0 auto; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
                    h2 { color: #e67e22; }
                </style>
            </head>
            <body>
                <div class="message">
                    <h2>ℹ️ Nothing to Update</h2>
                    <p>No analysis or corrective actions were provided.</p>
                    <a href="/" style="color: #0078D7; text-decoration: none; font-weight: 600;">← Back to Home</a>
                </div>
            </body>
            </html>
            ''', 200

        conn = get_db_connection()
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
            logger.info(f"✅ Successfully updated {len(kpi_data)} KPI value(s)")

            return f'''
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Submission Successful</title>
                <style>
                    body {{
                        font-family:'Segoe UI',sans-serif; background:#f4f4f4;
                        display:flex; justify-content:center; align-items:center;
                        min-height:100vh; margin:0; padding: 20px;
                    }}
                    .success-container {{
                        background:#fff; padding:50px; border-radius:10px;
                        box-shadow:0 4px 15px rgba(0,0,0,0.1); text-align:center;
                        max-width: 550px; width: 100%;
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
                        <a href="/form?responsible_id={responsible_id}&week={week}">🔄 View Form Again</a>
                        <a href="/dashboard" class="dashboard">📊 View Dashboard</a>
                    </div>
                </div>
            </body>
            </html>
            '''
        except Exception as e:
            if conn:
                conn.rollback()
            raise
        finally:
            return_db_connection(conn)

    except Exception as e:
        logger.error(f"❌ Error in submit_form: {str(e)}")
        traceback.print_exc()
        return f'<h2 style="color:red; padding: 20px;">❌ Failed to submit KPI values</h2><p>{str(e)}</p>', 500

@app.route('/trigger-task')
def trigger_task():
    """Manual trigger for scheduled task - useful for testing"""
    try:
        scheduled_email_task()
        return {
            'status': 'success',
            'message': 'Scheduled task executed manually',
            'timestamp': datetime.now(pytz.timezone('Africa/Tunis')).isoformat()
        }
    except Exception as e:
        logger.error(f"Error triggering task: {str(e)}")
        return {
            'status': 'error',
            'message': str(e)
        }, 500

# ========================================
# ERROR HANDLERS
# ========================================

@app.errorhandler(404)
def not_found(e):
    return '''
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Page Not Found</title>
        <style>
            body { font-family: 'Segoe UI', sans-serif; background: #f4f6f9; padding: 40px; text-align: center; }
            .error-container { background: white; padding: 50px; border-radius: 10px; max-width: 500px; margin: 0 auto; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
            h1 { color: #e74c3c; }
            a { color: #0078D7; text-decoration: none; font-weight: 600; }
        </style>
    </head>
    <body>
        <div class="error-container">
            <h1>404 - Page Not Found</h1>
            <p>The page you're looking for doesn't exist.</p>
            <a href="/">← Back to Home</a>
        </div>
    </body>
    </html>
    ''', 404

@app.errorhandler(500)
def internal_error(e):
    logger.error(f"Internal server error: {str(e)}")
    return '''
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Server Error</title>
        <style>
            body { font-family: 'Segoe UI', sans-serif; background: #f4f6f9; padding: 40px; text-align: center; }
            .error-container { background: white; padding: 50px; border-radius: 10px; max-width: 500px; margin: 0 auto; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
            h1 { color: #e74c3c; }
            a { color: #0078D7; text-decoration: none; font-weight: 600; }
        </style>
    </head>
    <body>
        <div class="error-container">
            <h1>500 - Server Error</h1>
            <p>Something went wrong. Please try again later.</p>
            <a href="/">← Back to Home</a>
        </div>
    </body>
    </html>
    ''', 500

# ========================================
# INITIALIZE SCHEDULER
# ========================================

scheduler = BackgroundScheduler(timezone=pytz.timezone('Africa/Tunis'))
scheduler.add_job(
    scheduled_email_task,
    'cron',
    hour=9,
    minute=30,
    timezone=pytz.timezone('Africa/Tunis'),
    id='kpi_email_scheduler',
    name='KPI Automated Email Scheduler'
)
scheduler.start()

logger.info("\n" + "="*70)
logger.info("✅ KPI AUTOMATION SYSTEM INITIALIZED")
logger.info("="*70)
logger.info(f"📅 Scheduler: Active")
logger.info(f"⏰ Schedule: Daily at 16:00 (Africa/Tunis)")
logger.info(f"📧 Next run: {scheduler.get_jobs()[0].next_run_time}")
logger.info(f"🌐 Server: Running on port {PORT}")
logger.info(f"🔗 Base URL: {_base_url()}")
logger.info("="*70 + "\n")

# ========================================
# START SERVER
# ========================================

if __name__ == '__main__':
    try:
        logger.info(f"🔗 Access points:")
        logger.info(f"   - Home: {_base_url()}/")
        logger.info(f"   - Dashboard: {_base_url()}/dashboard")
        logger.info(f"   - Health Check: {_base_url()}/health")
        logger.info("\n" + "="*70 + "\n")
        
        # Use 0.0.0.0 for Azure, debug=False for production
        app.run(host='0.0.0.0', port=PORT, debug=False, threaded=True)
    except KeyboardInterrupt:
        logger.info("\n🛑 Server stopped by user")
    finally:
        if db_pool:
            db_pool.closeall()
        scheduler.shutdown()
        logger.info("✅ Cleanup complete")
