#!/usr/bin/env python3
"""Streamlit entry point for the existing DOCX-to-PPT pipeline."""

from __future__ import annotations

import subprocess
import sys
import tempfile
from pathlib import Path

import streamlit as st


PROJECT_ROOT = Path(__file__).resolve().parent

st.set_page_config(page_title="广发策略PPT自动生成器", page_icon="📊")
st.title("广发策略PPT自动生成器")

uploaded_file = st.file_uploader("上传一个 DOCX 文件", type=["docx"])

if uploaded_file is not None:
    st.caption(f"已选择：{uploaded_file.name}")

    if st.button("生成 PPT", type="primary"):
        st.session_state.pop("generated_pptx", None)
        st.session_state.pop("generated_name", None)

        with st.status("正在生成 PPT，请稍候……", expanded=True) as status:
            try:
                with tempfile.TemporaryDirectory(prefix="ppt_generator_") as temp_dir:
                    temp_path = Path(temp_dir)
                    input_path = temp_path / "input.docx"
                    output_path = temp_path / f"{Path(uploaded_file.name).stem}.pptx"
                    input_path.write_bytes(uploaded_file.getvalue())

                    result = subprocess.run(
                        [
                            sys.executable,
                            str(PROJECT_ROOT / "test_generate.py"),
                            "--input",
                            str(input_path),
                            "--output",
                            str(output_path),
                            "--work-dir",
                            str(temp_path / "work"),
                        ],
                        cwd=PROJECT_ROOT,
                        capture_output=True,
                        text=True,
                        check=False,
                    )

                    if result.returncode != 0:
                        raise RuntimeError(
                            result.stderr.strip()
                            or result.stdout.strip()
                            or "PPT 生成失败"
                        )

                    st.session_state.generated_pptx = output_path.read_bytes()
                    st.session_state.generated_name = output_path.name
                    status.update(label="PPT 生成完成", state="complete")
            except Exception as exc:
                status.update(label="PPT 生成失败", state="error")
                st.error(str(exc))

if st.session_state.get("generated_pptx"):
    st.download_button(
        "下载生成的 PPT",
        data=st.session_state.generated_pptx,
        file_name=st.session_state.generated_name,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary",
    )
