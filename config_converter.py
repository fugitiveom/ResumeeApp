''' converting config.py into config.yaml '''
import yaml
import config

config_data = {
    'country': config.COUNTRY,
    'job_portal': config.JOB_PORTAL,
    'root_dir': config.ROOT_DIR,
    'job_type': config.JOB_TYPE,
    'email_regexp': config.EMAIL_REGEXP,
    'resume_regexp': config.RESUME_REGEXP,
    'cover_letter_regexp': config.COVER_LETTER_REGEXP,
    'ph_position_title': config.PH_POSITION_TITLE,
    'ph_company_name': config.PH_COMPANY_NAME,
    'ph_platform_source': config.PH_PLATFORM_SOURCE,
    'ph_date': config.PH_DATE,
    'resume_types': config.resume_types
}

with open('config.yaml', 'w', encoding='utf8') as yaml_file:
    yaml.dump(config_data, yaml_file)