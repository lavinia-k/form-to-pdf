FROM python:3
COPY main.py /
COPY utils.py /
COPY requirements.txt /
COPY .secrets /.secrets
RUN pip install -r requirements.txt
CMD [ "python", "./main.py" ]