from models import TaskChecker, LogDb, session, config_class


def task_checker_db(status, task_details, comment='', task_name='Get EMSX Trade', task_type='Task Scheduler '):
    if comment != '':
        comment_db = comment
    else:
        comment_db = 'Success'

    new_task_checker = TaskChecker(
        task_name=task_name,
        task_details=task_details,
        task_type=task_type,
        status=status,
        comment=comment_db
    )
    session.add(new_task_checker)
    session.commit()


def add_log_db(project, task, issue, description, msg_type):
    new_log_db = LogDb(
        project=project,
        task=task,
        issue=issue,
        description=description,
        msg_type=msg_type
    )
    session.add(new_log_db)
    session.commit()