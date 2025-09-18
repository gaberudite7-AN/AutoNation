# Dockerfile, Image, Container
FROM python

ADD UrbanScience_Update.py .

RUN pip install selenium undetected_chromedriver setuptools pandas numpy xlwings

CMD [ "python", "./UrbanScience_Update.py"]