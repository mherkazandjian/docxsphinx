FROM ubuntu:xenial

RUN \
  apt-get update && \
  apt-get -y install \
    software-properties-common \
    python3-pip \
    git

RUN \
  pip3 install \
    docutils==0.15 \
    sphinx==1.6.2 \
    python-docx==0.8.6 \
    sphinx-bootstrap-theme==0.6.4 \
    sphinxcontrib-websupport==1.0.1 \
    git+https://github.com/mherkazandjian/docxsphinx.git@master

RUN apt-get clean

ENTRYPOINT ["make", "docx", "html"]
