from database.conn import DBManager, func, literal_column, case, distinct, cast, text, desc, JSONB
from model.models import File, Job
from ppt_generator.ppt_table import ppt
from collections import Counter
from datetime import date, timedelta, datetime
import os
import logging

logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

try:
    db = DBManager()
    session = db.session
    logging.info("DBManager initialized successfully. Session is available.")
except Exception as e:
    logging.error("Initialization failed:", e)

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

        logging.info(f"Fetched and processed {len(results)} statuses excluding {exclude_result}")
        return [{"status": status, "count": count} for status, count in counts.items()]
    

    except Exception as e:
        logging.error("error in fetching exception",exc_info=True)
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
        logging.info(f"Successfully Fetched status from files")
        return [{"Title": title, "Count": func()} for title, func in metric_functions]

    except Exception as e:
        logging.error("Error in fetch_status_files:", exc_info=True)
        return []
    
def duplicates_from_source_category():
    date_start = '2025-04-01 00:00:00'
    date_end   = '2025-06-30 23:59:59'
    try:
        results = (
            session.query(
                File.meta_data['sourceCategory'].astext.label('source_category'),
                func.count(File.md5).label('count')
            )
            .filter(
                File.date_created >= date_start,
                File.date_created <= date_end,
                File.status == 'DUPLICATE'
            )
            .group_by(File.meta_data['sourceCategory'].astext)
            .all()
        )

        data = [{"Duplicates": title, "Count": count} for title, count in results]
        total = sum(item["Count"] for item in data)

        # Optionally append or return separately
        data.append({"Duplicates": "TOTAL", "Count": total})

        logging.info(f"Successfully fetched duplicates from source_category")
        return data

    except Exception as e:
        logging.error("error in fetching duplicates from source_category",e)
        return []
    
def processed_from_source_category():
    date_start = '2025-04-01 00:00:00'
    date_end   = '2025-06-30 23:59:59'
    try:
        results = (
            session.query(
                File.meta_data['sourceCategory'].astext.label('source_category'),
                func.count(File.md5).label('count')
            )
            .filter(
                File.date_created >= date_start,
                File.date_created <= date_end,
                File.status == 'PROCESSED'
            )
            .group_by(File.meta_data['sourceCategory'].astext)
            .all()
        )

        data = [{"Processed": title, "Count": count} for title, count in results]
        total = sum(item["Count"] for item in data)

        # Optionally append or return separately
        data.append({"Processed": "TOTAL", "Count": total})

        logging.info(f"Successfully fetched processed filed from source_category")
        return data

    except Exception as e:
        logging.error("error in fetching processed files from source_category",e)
        return []
    
def sourceCategory_count():
    try :
        subq = (
            session.query(
                File.meta_data['sourceCategory'].astext.label("source_category"),
                func.count(File.md5).label("count")
            )
            .group_by(File.meta_data['sourceCategory'].astext)
        ).subquery()

        # CASE logic for grouping categories with count < 1000
        group_case = case(
            (subq.c.count < 1000, literal_column("'Other Source Category < 1000 each'")),
            else_=subq.c.source_category
        ).label("source_category")

        # Main query: sum counts, group by adjusted category, order with 'Data less than 1k' at bottom
        query = (
            session.query(
                group_case,
                func.sum(subq.c.count).label("count")
            )
            .group_by(group_case)
            .order_by(
                case(
                    (group_case == 'Other Source Category < 1000 each', 1),
                    else_=0
                ),
                func.sum(subq.c.count).desc()
            )
        )


        # Return result as list of dictionaries
        results =  [{"SourceCategory": title, "Job Count": count} for title, count in query.all()]
        logging.info(f"Fetched {len(results)} grouped sourceCategory results")
        return results
    
    except Exception as e:
        logging.error("error in sourceCategory_count",e)
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

        job_done = sum(row.job_count for row in query.filter(Job.status_id == 5,
                                   File.s3_location.isnot(None)))
        job_done_within_SLA = sum(row.job_count for row in query.filter(Job.status_id == 5,
                                    File.s3_location.isnot(None),
                                    Job.last_modified_date > Job.submission_deadline))
        
        logging.info(f"Successfully fetched SLA Jobs")
        return [[{"Priority": priority, "SLA(hrs)": hours, "Job Count": jobCount} for priority, hours, jobCount in total_job_with_SLA], [{'job_done': job_done,'job_done_within_SLA': job_done_within_SLA}]]
    
    except Exception as e:
        logging.error("error in fetch_SLA_jobs",e)
        return []
    
def fetch_total_and_cancelled_jobs():
    curr_date = date.today()
    past_week = curr_date - timedelta(days = 7)
    counter = 0

    try: 
        job_tup = []
        total_job_count = (
            session
            .query(func.count(distinct(Job.job_id))
            .label("job_count"))
            .join(File, Job.job_id == File.job_id)
        )
        while counter < 8:

            job_date = datetime.strptime(f'{curr_date}', '%Y-%m-%d')
            job_past_week = datetime.strptime(f'{past_week}', '%Y-%m-%d')
            date_curr_str = f'{job_date.strftime("%b")} {job_date.day}'
            date_past_str = f'{job_past_week.strftime("%b")} {job_past_week.day}' 
            temp_tup = (f'{date_past_str} - {date_curr_str}',
                           total_job_count.filter(Job.date_created >= f"{past_week} 00:00:00",
                                                  Job.date_created <= f"{curr_date} 23:59:59")
                                          .scalar(),
                           total_job_count.filter(Job.status_id == 7,
                                                  Job.date_created >= f"{past_week} 00:00:00",
                                                  Job.date_created <= f"{curr_date} 23:59:59")
                                          .scalar())
            job_tup.append(temp_tup)
            curr_date = past_week
            past_week = past_week - timedelta(days = 7)
            counter += 1

        logging.info(f"Successfully fetch_total_and_cancelled_jobs")
        return [{"day_date": day_date, "TOTAL": job_count, "CANCELLED": cancelled_job} for day_date,job_count,cancelled_job in job_tup]
    except Exception as e:
        logging.error("error in fetch_total_and_cancelled_jobs",e)
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

        logging.info(f"Successfully fetch_jobs_by_source_category")
        return [{'Sources': sources, 'Jobs': jobs} for sources, jobs in sorted_over_1000.items()]

    
    except Exception as e:
        logging.error("error in fetch_jobs_by_source_category",e)
        return []

def generate_ppt(prs, 
                 title=None, 
                 table_data=None, 
                 table_data2=None, 
                 table_graph=None):
    logging.info(f"Adding slide for {title}")
    prs.add_slide()
    if title:
        prs.add_title(title)
    if table_data and title != 'Jobs by Priority':
        prs.add_table(table_data)
    if title == "Job Received Count":
        prs.add_slide()
        prs.add_title(title)
        prs.jobs_cancelled_add_graph(table_graph)
    elif title == "Jobs by Priority":
        prs.add_SLA_table(table_data)
        prs.add_slide()
        prs.add_title(title)
        prs.add_SLA_graph(table_graph)
    elif table_graph:
        prs.add_slide()
        if title:
            prs.add_title(title)
        prs.add_graph(table_graph)
    prs.save()
    os.startfile(path)

if __name__ == "__main__":
    path = "status_report.pptx"
    exclude_result = ['DONE', 'PROCESSING','UNKNOWN','DUPLICATE','PROCESSED']

    # Table data
    exception_result = fetch_exception(exclude_result)
    status_data = fetch_status_files()
    duplicate_status = duplicates_from_source_category()
    processed_status = processed_from_source_category()
    source_category_summary = sourceCategory_count()
    job_done_with_SLA = fetch_SLA_jobs()
    job_per_source = fetch_jobs_by_source_category()
    # total_job_count = fetch_total_and_cancelled_jobs()

    logging.info("Generating Powerpoint")
    prs = ppt(path)
    generate_ppt(prs, 
        title="Exceptions Encountered in Jobs Processing", 
        table_data=exception_result, 
        table_data2=None, 
        table_graph=None)
    
    generate_ppt(prs, 
        title="Duplicate by Hash", 
        table_data=status_data, 
        table_data2=None, 
        table_graph=None)
    
    generate_ppt(prs, 
        title="Deduped vs Processed", 
        table_data=duplicate_status, 
        table_data2=processed_status, 
        table_graph=None)

    generate_ppt(prs, 
        title="Source Category Summary", 
        table_data=source_category_summary, 
        table_data2=None, 
        table_graph=None)
    
    generate_ppt(prs, 
        title="Jobs by Priority", 
        table_data=job_done_with_SLA, 
        table_data2=None,
        table_graph=job_done_with_SLA)
    
    # generate_ppt(prs, 
    #     title="Jobs by Source Category", 
    #     table_data=job_per_source, 
    #     table_data2=None,
    #     table_graph=job_per_source)
    
    # generate_ppt(prs, 
    #     title="Job Received Count", 
    #     table_data=total_job_count, 
    #     table_data2=None,
    #     table_graph=total_job_count)

    logging.info("Powerpoint Generated")