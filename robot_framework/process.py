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
    Fetches data from the DLPGoogleAlertsView, sends an email notification to the user,
    and updates the isNotified status in the database.

    Args:
        orchestrator_connection (OrchestratorConnection): An instance of OrchestratorConnection
                                                          to get the database connection string
                                                          and other constants.

    Raises:
        pyodbc.Error: If there is an error in the database connection or operations.
        ValueError: If there is an issue with the input data.
    """
    try:
        connection_string = orchestrator_connection.get_constant('DbConnectionString').value
        email_sender = orchestrator_connection.get_constant("e-mail_noreply").value
        oc_args_json = json.loads(orchestrator_connection.process_arguments)
        subject = oc_args_json['subject']
        body_template = oc_args_json['body']

        with pyodbc.connect(connection_string) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT [alertId], [triggerUserEmail], [link], [isNotified], [azident], [navn], [email_ad]
                FROM [RPA].[rpa].[DLPGoogleAlertsView]
                WHERE isNotified = 0
                      AND triggerType = 'CPR-Number'
                      AND email_ad IS NOT NULL
            """)

            rows = cursor.fetchall()
            for row in rows:
                to_email = row.email_ad
                to_name = row.navn
                link_to_file = row.link

                body = body_template.format(to_name=to_name, link_to_file=link_to_file)

                smtp_util.send_email(
                    receiver=to_email,
                    sender=email_sender,
                    subject=subject,
                    body=body,
                    html_body=True,
                    smtp_server=config.SMTP_SERVER,
                    smtp_port=config.SMTP_PORT
                )

                cursor.execute("EXEC [rpa].[DLPGoogleAlerts_Insert] @alertId = ?, @isNotified = ?", row.alertId, 1)
                conn.commit()

    except pyodbc.Error as e:
        print(f"Database error: {str(e)}")
    except json.JSONDecodeError as e:
        print(f"JSON decode error: {str(e)}")
    except KeyError as e:
        print(f"Missing key in process arguments: {str(e)}")
    except ValueError as e:
        print(f"Value error: {str(e)}")


if __name__ == "__main__":
    oc = OrchestratorConnection.create_connection_from_args()
    process(oc)
