FROM python:3.11

WORKDIR /etl_test

COPY requirements.txt .
COPY . ./etl_test/

RUN pip install -r requirements.txt
CMD [ "python", ".etl_test/main_pd.py" ]
