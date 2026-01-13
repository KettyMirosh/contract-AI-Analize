import pytest
from app import app, allowed_file

@pytest.fixture
def client():
    app.config['TESTING'] = True
    with app.test_client() as client:
        yield client

def test_index_page(client):
    response = client.get('/')
    assert response.status_code == 200
    assert b'AI Contract Analyzer' in response.data

def test_allowed_file():
    assert allowed_file('test.docx') == True
    assert allowed_file('test.doc') == False
    assert allowed_file('test.pdf') == False
    assert allowed_file('test.txt') == False
    assert allowed_file('document.DOCX') == True

def test_upload_no_file(client):
    response = client.post('/upload')
    assert response.status_code == 400

def test_upload_empty_filename(client):
    data = {'contract': (None, '')}
    response = client.post('/upload', data=data)
    assert response.status_code == 400
