from database.conn import DBManager, func
from model.models import File
from ppt_generator.ppt_table import ppt
from collections import Counter

db = DBManager()
session = db.session

def fetch_exception(exclude_result):
    try:
        results = (
            session.query(File.status)
            .filter(~File.status.in_(exclude_result))
            .all()
        )

        # Strip and count manually
        cleaned = [status[0].strip() for status in results]
        counts = Counter(cleaned)

        return [{"status": status, "count": count} for status, count in counts.items()]

    except Exception as e:
        print(e)
        return []
    
def fetch_status_files():
    try:
        # Metric functions with label and callable logic
        metric_functions = [
            (
                "Total Files", lambda: session
                .query(func.count())
                .select_from(File)
                .scalar()),
            (
                "Processed Files", lambda: session
                .query(func.count())
                .filter(File.status != 'PROCESSING')
                .scalar()),
            (
                "Deduplicated Files", lambda: 
                sum(row.duplicates for row in session
                    .query(File.md5, (func.count(File.id) - 1)
                    .label("duplicates"))
                    .group_by(File.md5)
                    .having(func.count(File.id) > 1)
                    .all())),
            (
                "Duplicate Groups", lambda: 
                len(session
                    .query(File.md5)
                    .group_by(File.md5)
                    .having(func.count(File.id) > 1)
                    .all())),
            (
                "Unique Files", lambda: 
                sum(row.unique for row in session
                    .query(File.md5, func.count(File.id)
                    .label("unique"))
                    .group_by(File.md5)
                    .having(func.count(File.id) == 1)
                    .all())),
            (
                "Null Files", lambda: session
                .query(func.count())
                .filter(File.status.is_(None))
                .scalar()),
        ]

        # Evaluate and return all metrics dynamically
        return [{"Title": title, "Count": func()} for title, func in metric_functions]

    except Exception as e:
        print("Error in fetch_status_files:", e)
        return []
    
def generate_ppt(prs, title, table_data=None, table_graph=None):
    prs.add_slide()
    if title:
        prs.add_title(title)
    if table_data:
        prs.add_table(table_data)
    if table_graph:
        prs.add_slide()
        prs.add_title(title)
        prs.add_graph(table_graph)
    prs.save()

if __name__ == "__main__":
    exclude_result = ['DONE', 'PROCESSING','UNKNOWN']
    exception_result = fetch_exception(exclude_result)
    status_data = fetch_status_files()

    prs = ppt("status_report.pptx")
    generate_ppt(prs, "Exceptions Encountered in Jobs Processing", exception_result, exception_result)
    generate_ppt(prs, "Deduped vs Processed Files", status_data, status_data)
    prs.save()

    
