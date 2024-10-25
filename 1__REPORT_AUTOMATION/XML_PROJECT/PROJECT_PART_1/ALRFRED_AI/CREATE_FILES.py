import os

# List of your apps, including helper_funcs and configs
apps = [
    "allegro_app", "apim_app", "belvis_app", "data_mart_app", "datalake_yin_app",
    "dev_job_monitoring_app", "eex_app", "epex_app", "euler_prod_app", "euler_qa_app",
    "go_manager_app", "hedging_app", "pfc_app", "prod_job_monitoring_app", "siloveda_app",
    "swissgrid_app", "timesteps_prod_app", "timesteps_qa_app",
    "volue_algo_trader_app", "volueinsight_app", "helper_funcs", "configs"
]

# Template for the setup.py file
setup_template = """
from setuptools import setup, find_packages

setup(
    name='{app_name}',
    version='0.1.0',
    packages=find_packages(include=['{app_name}']),
    package_dir={{'{app_name}': '{app_name}'}},
    include_package_data=True,
    author='bh_dict',
    author_email='Verteiler.BHB@iwb-ch',
    scripts=['{app_name}/ingestions/*.py'],
    install_requires=[],
    description='{app_name} for data processing',
    entry_points={{
        'console_scripts': [
            'run_{app_name}={app_name}.main:main',
        ]
    }}
)
"""

# Loop through each app and create its setup.py file
for app in apps:
    setup_content = setup_template.format(app_name=app)
    setup_file_path = f'setup_{app}.py'
    
    with open(setup_file_path, 'w') as f:
        f.write(setup_content)

    print(f'Created setup.py for {app} at {setup_file_path}')