import gradio as gr

# Change this to the actual path in your Drive
mapFileName = 'assets/konta_KG.xlsx'

def process_excel(file_obj):
    # 1. Load the uploaded file
    inputFileName= file_obj.name

    df = loadSAPData(inputFileName, mapFileName)
    output_path = writeNCNReport(df, inputFileName)

    return output_path

# 4. Create the Web UI
client = gr.Interface(
    fn=process_excel,
    inputs=gr.File(label="Upload end of the year SAP spreadsheet"),
    outputs=gr.File(label="Download spreadsheet with NCN report tables."),
    title="SAP -> NCN reformatter. Use at your own risk.",
    description="Upload a SAP report you get at the end of the year to get a table with NCN categories."
)

client.launch(share=True)