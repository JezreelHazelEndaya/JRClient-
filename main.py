from database.conn import DBManager, func, distinct, JSONB, cast, text, desc
from model.models import File, Job, Status
from ppt_generator.ppt_table import ppt
from collections import Counter

db = DBManager()
session = db.session

def fetch_total_and_cancelled_jobs():
    try: 
        total_job_count = (
            session
            .query(func.count(distinct(Job.job_id))
            .label("job_count"))
            .join(File, Job.job_id == File.job_id)
        )
        job_dict = [('Job Count', total_job_count.scalar()), ('Cancelled Job Count', total_job_count.filter(Job.status_id == 7).scalar())]
        return [{'title': title, 'count': count} for title, count in job_dict]
    except Exception as e:
        print(e)
        return []
    
def fetch_SLA_jobs():
    try:
        query = (
            session.query(
                Job.message_priority,
                func.count(Job.job_id).label("job_count")
            )
            .join(File, Job.job_id == File.job_id)
            .group_by(Job.message_priority)
            .order_by(desc(Job.message_priority))
        )

        total_job_with_SLA = []
        for item in query.all():
            if item[0] == 7:
                new_item = (item[0], "12hrs", item[1])
            elif  item[0] == 6:
                new_item = (item[0], "24hrs", item[1])
            elif  item[0] == 5:
                new_item = (item[0], "36hrs", item[1])
            elif  item[0] == 4:
                new_item = (item[0], "48hrs", item[1])
            elif  item[0] == 3:
                new_item = (item[0], "60hrs", item[1])
            elif  item[0] == 2:
                new_item = (item[0], "72hrs", item[1])
            elif  item[0] == 1:
                new_item = (item[0], ">84hrs", item[1])
            else:
                new_item = item
            total_job_with_SLA.append(new_item)

        job_for_SLA = query.filter(Job.status_id == 5,
                                   File.s3_location.isnot(None))
        job_within_SLA = query.filter(Job.status_id == 5,
                                    File.s3_location.isnot(None),
                                    Job.last_modified_date > Job.submission_deadline)
        return [{"Priority": priority, "SLA(hrs)": hours, "Job Count": jobCount} for priority, hours, jobCount in total_job_with_SLA]
    
    except Exception as e:
        print(e)
        return []
    
def fetch_jobs_by_source_category():
    try:
        source_category = File.meta_data.cast(JSONB)['sourceCategory'].astext.label("source_category")

        query = (
            session.query(
                source_category,
                func.count(Job.job_id).label("job_count")
            )
            .join(File, Job.job_id == File.job_id)
            .group_by(source_category)
        )

        source_counts = query.all()
        over_1000 = {}
        sum_under_1000 = 0

        for source, count in source_counts:
            if source is None or source.strip() == "":
                source = "N/A"

            if count >= 1000:
                over_1000[source] = count
            elif count < 1000:
                sum_under_1000 += count


        sorted_over_1000 = dict(sorted(over_1000.items(), key=lambda item: item[1], reverse=True))
        sorted_over_1000["Sources w/ Job <1000"] = sum_under_1000

        return [{'Sources': sources, 'Jobs': jobs} for sources, jobs in sorted_over_1000.items()]

    
    except Exception as e:
        print(e)
        return []
    
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
    total_job_count = fetch_total_and_cancelled_jobs()
    job_per_source = fetch_jobs_by_source_category()
    job_done_with_SLA = fetch_SLA_jobs()



    prs = ppt("status_report.pptx")
    generate_ppt(prs, "Total & Cancelled Jobs", total_job_count, total_job_count)
    generate_ppt(prs, "Exceptions Encountered in Jobs Processing", exception_result, exception_result)
    generate_ppt(prs, "Deduped vs Processed Files", status_data, status_data)
    generate_ppt(prs, "Jobs per Source Category", job_per_source, job_per_source)
    generate_ppt(prs, "Jobs by Priority", job_done_with_SLA)
    prs.save()
    


    
