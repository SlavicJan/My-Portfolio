#DAG FOR AIRFLOW
#FOR apache-airflow-providers-postgres

# Настройка рассылки отчёта в Apache Airflow

#1. Установи зависимости  
#   - `apache-airflow`  
#   - `apache-airflow-providers-postgres` (или нужный провайдер для твоей БД)  
#   - `apache-airflow-providers-smtp`  

#2. Создай новый DAG-файл, например `dags/abc_xyz_report.py`:

from airflow import DAG
from airflow.providers.postgres.operators.postgres import PostgresHook
from airflow.operators.email import EmailOperator
from airflow.operators.python import PythonOperator
from datetime import datetime

DEFAULT_ARGS = {
    'start_date': datetime(2025, 9, 9),
    'email_on_failure': False,
    'email_on_retry': False,
}

with DAG(
    dag_id='abc_xyz_weekly_report',
    default_args=DEFAULT_ARGS,
    schedule_interval='@weekly',
    catchup=False
) as dag:

    def fetch_report(**context):
        hook = PostgresHook(postgres_conn_id='my_postgres')
        sql = """
        -- (здесь вставь финальный SQL-запрос из блока выше)
        """
        df = hook.get_pandas_df(sql)
        html = df.to_html(index=False)
        context['ti'].xcom_push(key='report_html', value=html)

    prepare_report = PythonOperator(
        task_id='prepare_report',
        python_callable=fetch_report,
        provide_context=True
    )

    send_email = EmailOperator(
        task_id='send_report_email',
        to=['team@example.com'],
        subject='Weekly ABC+XYZ Report',
        html_content="{{ ti.xcom_pull(key='report_html') }}",
        smtp_conn_id='smtp_default'
    )

    prepare_report >> send_email


#3. Прописать соединения в Airflow UI  
#   - **Postgres Connection** (`Conn Id`: `my_postgres`)  
#     • Тип: Postgres  
#     • Host, Schema, Login, Password, Port  
#   - **SMTP Connection** (`Conn Id`: `smtp_default`)  
#     • Тип: SMTP  
#     • SMTP Host, Port, логин и пароль почты  

#4. Перезапусти Airflow Scheduler и Webserver in CMD:
#   systemctl restart airflow-scheduler
#   systemctl restart airflow-webserver

 #  После этого DAG `abc_xyz_weekly_report` появится в UI.  

#5. Проверка и мониторинг  
#   - Зайди в Web UI → DAGs → `abc_xyz_weekly_report`  
#   - Нажми **Trigger DAG** для тестового запуска  
#   - В логах убедись, что подготовка отчёта и отправка письма прошли без ошибок  


#Теперь отчёт по ABC+XYZ автоматически будет собираться раз в неделю и отправляться на указанный email.