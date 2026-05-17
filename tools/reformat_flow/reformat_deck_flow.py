from ibm_watsonx_orchestrate.agent_builder.tools import tool, WXOFile, MultiFileConstraints
from ibm_watsonx_orchestrate.agent_builder.connections import ConnectionType
from ibm_watsonx_orchestrate.run import connections

from typing import List, Annotated


# Import sub-modules locally
from extract_guidelines import extract_guidelines, build_model
from extract_pptx_content import extract_all_content
from render_deck import render_slides_preserve_structure

@tool(
    name="reformat_deck",
    description="Rebuilds presentation using structural mapping and style guidelines PDF.",
    expected_credentials=[{"app_id": "watsonx_model_creds", "type": ConnectionType.KEY_VALUE}]
)
def reformat_deck(
    files: Annotated[
        List[WXOFile],
        MultiFileConstraints(
            min_files=2,
            max_files=2,
            accepted_file_extensions=["pptx", "pdf"],
        ),
    ],
) -> bytes:
    
    #  1. Validate and read files into memory, ensuring we have exactly one PDF and one PPTX
    if len(files) != 2:
        raise ValueError(f"Expected exactly 2 files, got {len(files)}.")

    blobs = []
    for i, ref in enumerate(files):
        data = WXOFile.get_content(ref)

        if not isinstance(data, (bytes, bytearray)):
            raise TypeError(
                f"file[{i}]: WXOFile.get_content returned {type(data).__name__}, not bytes"
            )
        blobs.append(bytes(data))

    def kind(buf: bytes) -> str:
        if buf.startswith(b"%PDF"):
            return "pdf"
        if buf.startswith(b"PK\x03\x04"):
            return "pptx"
        return "unknown"

    kinds = [kind(b) for b in blobs]
    if sorted(kinds) != ["pdf", "pptx"]:
        raise ValueError(
            f"Need one distinct .pptx and one distinct .pdf. Got kinds={kinds}, "
            f"sizes={[len(b) for b in blobs]}. Please re-upload one of each."
        )

    pptx_bytes = blobs[kinds.index("pptx")]
    pdf_bytes = blobs[kinds.index("pdf")]

    # 2. Set up credentials for Watsonx API calls
    creds = connections.key_value("watsonx_model_creds")

    api_key = creds.get("api_key", None)
    project_id = creds.get("project_id", None)
    url = creds.get("watsonx_url", None)

    if not all([api_key, project_id, url]):
        raise ValueError(
            "Missing required credentials for Watsonx API. Ensure 'api_key', 'project_id', and 'watsonx_url' are all provided in the 'watsonx_model_creds' connection."
        )
    
    # 3. Extract guidelines from PDF
    text_model = build_model(api_key, project_id, url)
    vison_model = build_model(api_key, project_id, url, model_id="meta-llama/llama-3-2-11b-vision-instruct")

    guidelines = extract_guidelines(pdf_bytes, text_model, vison_model)

    # 4. Extract content from PPTX
    extracted_deck = extract_all_content(pptx_bytes)

    # 5. Render new PPTX using extracted content and guidelines
    output_bytes = render_slides_preserve_structure(extracted_deck, guidelines)
    return output_bytes