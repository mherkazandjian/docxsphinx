from docxsphinx.builder import DocxBuilder

def setup(app):
    app.add_builder(DocxBuilder)
    app.add_config_value('docx_template', None, 'env')
