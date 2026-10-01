import azure.functions as func
import logging
import sys
import os

# Dodanie katalogu nadrzędnego do ścieżki Pythona, aby użyć skryptu generate_calendar
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from generate_calendar import download_excel, parse_schedule, generate_ics, EXCEL_URL

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

@app.route(route="calendar", methods=["GET"])
def get_ics_calendar(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('HTTP trigger: zapytanie o kalendarz ICS.')
    
    group = req.params.get('group', '6')
    excel_path = "/tmp/plan_azure.xlsx"
    ics_path = f"/tmp/grupa_{group}.ics"
    
    try:
        download_excel(EXCEL_URL, excel_path)
        events = parse_schedule(excel_path, target_group=group)
        generate_ics(events, ics_path, title=f"Plan Zajęć UR - Grupa {group}")
        
        with open(ics_path, "r", encoding="utf-8") as f:
            ics_content = f.read()
            
        return func.HttpResponse(
            ics_content,
            mimetype="text/calendar",
            headers={
                "Content-Disposition": f"inline; filename=grupa_{group}.ics"
            },
            status_code=200
        )
    except Exception as e:
        logging.error(f"Błąd podczas generowania kalendarza: {e}")
        return func.HttpResponse(f"Błąd serwera: {str(e)}", status_code=500)
