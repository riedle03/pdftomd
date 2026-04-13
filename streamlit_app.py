"""Streamlit web app for document-to-Markdown conversion."""

from __future__ import annotations

import tempfile
import zipfile
import io
from pathlib import Path

import streamlit as st

from convert_to_md import (
    MARKITDOWN_EXTENSIONS,
    HWP_EXTENSIONS,
    PDF_EXTENSIONS,
    convert_one,
)

# 웹 배포에서 지원하는 확장자 (pywin32 제외)
WEB_SUPPORTED = PDF_EXTENSIONS | HWP_EXTENSIONS | MARKITDOWN_EXTENSIONS
ALLOWED_EXTENSIONS = sorted(ext.lstrip(".") for ext in WEB_SUPPORTED)

st.set_page_config(
    page_title="Document to Markdown Converter",
    page_icon="📄",
    layout="centered",
)

st.title("Document to Markdown Converter")
st.caption("PDF, HWP, Office(docx/xlsx/pptx), 이미지, CSV 등 → Markdown 변환")

# 옵션
extract_images = st.checkbox("PDF/HWP 이미지 추출", value=True)

# 파일 업로드 (복수)
uploaded_files = st.file_uploader(
    "변환할 파일을 업로드하세요",
    type=ALLOWED_EXTENSIONS,
    accept_multiple_files=True,
)

if uploaded_files and st.button("변환 시작", type="primary"):
    results: list[tuple[str, str, bytes]] = []  # (filename, md_filename, md_bytes)
    asset_files: list[tuple[str, bytes]] = []  # (archive_path, bytes)
    progress = st.progress(0, text="변환 준비 중...")

    with tempfile.TemporaryDirectory(prefix="streamlit_convert_") as tmp_dir:
        tmp_path = Path(tmp_dir)
        input_dir = tmp_path / "input"
        output_dir = tmp_path / "output"
        input_dir.mkdir()
        output_dir.mkdir()

        total = len(uploaded_files)
        success = 0
        failed = 0

        for i, uploaded in enumerate(uploaded_files):
            progress.progress((i) / total, text=f"변환 중... {i+1}/{total}: {uploaded.name}")

            # 업로드 파일을 임시 디렉토리에 저장
            src_file = input_dir / uploaded.name
            src_file.write_bytes(uploaded.getbuffer())

            ext = src_file.suffix.lower()
            if ext not in WEB_SUPPORTED:
                st.warning(f"지원하지 않는 형식: {uploaded.name} ({ext})")
                failed += 1
                continue

            try:
                result = convert_one(src_file, output_dir, extract_images=extract_images)
                md_content = result.output.read_bytes()
                results.append((uploaded.name, result.output.name, md_content))

                # 이미지 에셋이 있으면 수집
                asset_dir = output_dir / f"{src_file.stem}_assets"
                if asset_dir.exists():
                    for asset_file in asset_dir.iterdir():
                        if asset_file.is_file():
                            asset_files.append(
                                (f"{asset_dir.name}/{asset_file.name}", asset_file.read_bytes())
                            )

                success += 1
            except Exception as e:
                st.error(f"변환 실패: {uploaded.name} - {e}")
                failed += 1

        progress.progress(1.0, text="변환 완료!")

    # 결과 표시
    st.divider()
    st.subheader(f"결과: {success}개 성공, {failed}개 실패")

    if results:
        # 단일 파일 + 에셋 없음 → md만 다운로드
        if len(results) == 1 and not asset_files:
            name, md_name, md_bytes = results[0]
            st.download_button(
                label=f"다운로드: {md_name}",
                data=md_bytes,
                file_name=md_name,
                mime="text/markdown",
            )
            with st.expander("미리보기", expanded=True):
                st.markdown(md_bytes.decode("utf-8")[:3000])

        # 에셋이 있거나 복수 파일 → ZIP (md + 이미지 에셋 포함)
        else:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for _, md_name, md_bytes in results:
                    zf.writestr(md_name, md_bytes)
                for archive_path, data in asset_files:
                    zf.writestr(archive_path, data)
            zip_buffer.seek(0)

            asset_count = len(asset_files)
            label = f"전체 다운로드 (ZIP, {len(results)}개 파일"
            if asset_count:
                label += f" + {asset_count}개 이미지"
            label += ")"

            st.download_button(
                label=label,
                data=zip_buffer,
                file_name="converted_markdown.zip",
                mime="application/zip",
            )

            if asset_count:
                st.info(f"추출된 이미지 {asset_count}개가 ZIP에 포함되었습니다.")

            # 개별 미리보기
            for name, md_name, md_bytes in results:
                with st.expander(f"{md_name}"):
                    st.markdown(md_bytes.decode("utf-8")[:2000])

# 푸터
st.divider()
st.caption("made by 이대형 with Claude Code")
