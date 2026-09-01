#!/usr/bin/env python3
"""Streamlit entry point for the existing DOCX-to-PPT pipeline."""

from __future__ import annotations

import hashlib
import subprocess
import sys
import tempfile
from pathlib import Path

import streamlit as st


PROJECT_ROOT = Path(__file__).resolve().parent


def user_facing_pipeline_error(stderr: str, stdout: str) -> str:
    """Return a concise upload-facing message instead of a raw traceback."""

    details = (stderr or stdout or "").strip()
    for line in reversed(details.splitlines()):
        normalized = line.strip()
        if "图片“" in normalized and any(
            marker in normalized
            for marker in ("文件过大", "分辨率过高", "无法识别", "已损坏")
        ):
            return normalized.split(":", 1)[-1].strip()
    if "DecompressionBomb" in details:
        return "报告中存在超高分辨率图片，请在Word中压缩图片后重新上传。"
    return "PPT生成失败，请检查上传的DOCX文件；技术详情已记录在服务日志中。"


st.set_page_config(page_title="广发策略PPT自动生成器", page_icon="📊")
st.title("广发策略PPT自动生成器")

uploaded_file = st.file_uploader("上传一个 DOCX 文件", type=["docx"])
current_upload_hash: str | None = None

if uploaded_file is not None:
    uploaded_bytes = uploaded_file.getvalue()
    current_upload_hash = hashlib.sha256(uploaded_bytes).hexdigest()
    if st.session_state.get("uploaded_docx_hash") != current_upload_hash:
        st.session_state.pop("generated_pptx", None)
        st.session_state.pop("generated_name", None)
        st.session_state.pop("generated_input_hash", None)
        st.session_state.uploaded_docx_hash = current_upload_hash

    st.caption(f"已选择：{uploaded_file.name}")

    if st.button("生成 PPT", type="primary"):
        st.session_state.pop("generated_pptx", None)
        st.session_state.pop("generated_name", None)
        st.session_state.pop("generated_input_hash", None)

        with st.status("正在生成 PPT，请稍候……", expanded=True) as status:
            try:
                with tempfile.TemporaryDirectory(prefix="ppt_generator_") as temp_dir:
                    temp_path = Path(temp_dir)
                    input_path = temp_path / "input.docx"
                    output_path = temp_path / f"{Path(uploaded_file.name).stem}.pptx"
                    input_path.write_bytes(uploaded_bytes)

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
                        technical_details = (
                            result.stderr.strip()
                            or result.stdout.strip()
                            or "PPT 生成失败"
                        )
                        print(technical_details, file=sys.stderr)
                        raise RuntimeError(
                            user_facing_pipeline_error(
                                result.stderr,
                                result.stdout,
                            )
                        )

                    st.session_state.generated_pptx = output_path.read_bytes()
                    st.session_state.generated_name = output_path.name
                    st.session_state.generated_input_hash = current_upload_hash
                    status.update(label="PPT 生成完成", state="complete")
            except Exception as exc:
                status.update(label="PPT 生成失败", state="error")
                st.error(str(exc))

if (
    current_upload_hash is not None
    and st.session_state.get("generated_pptx")
    and st.session_state.get("generated_input_hash") == current_upload_hash
):
    st.download_button(
        "下载生成的 PPT",
        data=st.session_state.generated_pptx,
        file_name=st.session_state.generated_name,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary",
    )
