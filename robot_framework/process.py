"""This module contains the main process of the robot."""
import json
import pyodbc

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

from itk_dev_shared_components.smtp import smtp_util

from robot_framework import config


def process(orchestrator_connection: OrchestratorConnection) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")
    fetch_data_and_send_emails(orchestrator_connection)


def fetch_data_and_send_emails(orchestrator_connection: OrchestratorConnection):
    """
    test
    """
    try:
        with pyodbc.connect(orchestrator_connection.get_constant('DbConnectionString').value) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT TOP 1 [alertId], [triggerUserEmail], [link], [isNotified], [azident], [navn], [email_ad]
                FROM [RPA].[rpa].[DLPGoogleAlertsView]
                WHERE isNotified = 0
                      AND triggerType = 'CPR-Number'
                      AND alertId = '95EBECFD-AE06-4906-89B4-F9DF0EAAE128'
                """)

            rows = cursor.fetchall()
            oc_args_json = json.loads(orchestrator_connection.process_arguments)
            for row in rows:
                to_email = row.email_ad
                to_name = row.navn
                link_to_file = row.link
                subject = oc_args_json['subject']
                body = oc_args_json['body']

                smtp_util.send_email(receiver=to_email, sender=orchestrator_connection.get_constant("E-mail").value, subject=subject, body=body, html_body=True, smtp_server=config.SMTP_SERVER, smtp_port=config.SMTP_PORT)

                cursor.execute("EXEC [rpa].[DLPGoogleAlerts_Insert] @alertId = ?, @isNotified = ?", row.alertId, 1)
                conn.commit()
    except pyodbc.Error as e:
        print(f"Database error: {str(e)}")
    except Exception as e:
        print(f"An unexpected error occurred: {str(e)}")


if __name__ == "__main__":
    oc = OrchestratorConnection.create_connection_from_args()
    process(oc)
