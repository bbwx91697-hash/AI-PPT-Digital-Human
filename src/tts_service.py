"""Multi-engine TTS service: Qwen3-TTS local GPU, Edge Neural, Windows SAPI, and Gemini 3.8."""

from __future__ import annotations

import asyncio
import atexit
from collections import deque
import json
import queue
import subprocess
import tempfile
import threading
from pathlib import Path

try:
    import edge_tts
except ImportError:
    edge_tts = None

try:
    import audio_manager
    import gemini_tts
except ImportError:
    from src import audio_manager, gemini_tts

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_KEY_PATH = ROOT / "config" / "gemini_key.dpapi"
CREDENTIAL_MODULE = ROOT / "src" / "credential_store.psm1"
QWEN3_PYTHON = ROOT / "venv_qwen_tts" / "Scripts" / "python.exe"
QWEN3_MODEL_DIR = ROOT / "models" / "qwen3_tts" / "Qwen3-TTS-12Hz-0.6B-CustomVoice"
QWEN3_WORKER = ROOT / "src" / "qwen3_worker.py"

MODELS = gemini_tts.MODELS
AVAILABLE_VOICES = ("Leda", "Achird", "Aoede", "Kore", "Puck")

ENGINE_LABELS = {
    "qwen3": "Qwen3-TTS 本地高音质 (推荐·免额度)",
    "edge": "微软神经语音 (免费·免Key)",
    "windows": "Windows 离线语音 (断网可用)",
    "gemini": "Gemini 3.8 云端语音",
}
LABEL_TO_ENGINE = {v: k for k, v in ENGINE_LABELS.items()}

QWEN3_VOICES = (
    ("serena", "Serena · 温和女声 (推荐)"),
    ("vivian", "Vivian · 明亮女声"),
    ("dylan", "Dylan · 清晰男声"),
    ("eric", "Eric · 青年男声"),
    ("uncle_fu", "Uncle Fu · 沉稳男声"),
    ("ryan", "Ryan · 英文男声"),
    ("aiden", "Aiden · 英文青年男声"),
    ("ono_anna", "Ono Anna · 日语女声"),
    ("sohee", "Sohee · 韩语女声"),
)
QWEN3_VOICE_CODES = tuple(code for code, _ in QWEN3_VOICES)

EDGE_VOICES = (
    ("zh-CN-XiaoxiaoNeural", "晓晓 · 温和女声 (推荐)"),
    ("zh-CN-XiaoyiNeural", "晓伊 · 清晰女声"),
    ("zh-CN-YunxiNeural", "云希 · 自然男声"),
    ("zh-CN-YunjianNeural", "云健 · 沉稳男声"),
    ("zh-CN-YunyangNeural", "云扬 · 播音男声"),
)
EDGE_VOICE_CODES = tuple(code for code, _ in EDGE_VOICES)

WINDOWS_VOICES = ("Microsoft Huihui Desktop", "Microsoft Zira Desktop")


def _run_powershell(script: str, stdin_text: str | None = None) -> str:
    completed = subprocess.run(
        ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        input=stdin_text,
        text=True,
        encoding="utf-8",
        capture_output=True,
        check=False,
    )
    if completed.returncode != 0:
        err = (completed.stderr or completed.stdout or "PowerShell 执行失败").strip()
        raise RuntimeError(err)
    return completed.stdout


def has_api_key(key_path: Path = DEFAULT_KEY_PATH) -> bool:
    return Path(key_path).is_file() and Path(key_path).stat().st_size > 0


def save_api_key(plain_key: str, key_path: Path = DEFAULT_KEY_PATH) -> Path:
    key = plain_key.strip()
    if not key:
        raise ValueError("请输入有效的 Gemini API Key。")
    target = Path(key_path).resolve()
    ps = (
        f"Import-Module '{CREDENTIAL_MODULE}' -Force; "
        f"$plain = [Console]::In.ReadToEnd().Trim(); "
        f"$sec = ConvertTo-SecureString $plain -AsPlainText -Force; "
        f"Save-GeminiKey -Key $sec -Path '{target}'"
    )
    _run_powershell(ps, stdin_text=key)
    return target


def load_api_key(key_path: Path = DEFAULT_KEY_PATH) -> str:
    target = Path(key_path).resolve()
    if not target.is_file():
        raise FileNotFoundError("尚未配置 Gemini API Key，请先在管理员设置中保存。")
    ps = (
        f"Import-Module '{CREDENTIAL_MODULE}' -Force; "
        f"$sec = Get-GeminiKey -Path '{target}'; "
        f"$ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec); "
        f"try {{ [Console]::Out.Write([Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)) }} "
        f"finally {{ [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) }}"
    )
    plain = _run_powershell(ps).strip()
    if not plain:
        raise RuntimeError("无法解密本机 Gemini API Key。")
    return plain


def _edge_rate(pace: str) -> str:
    mapping = {"偏慢": "-15%", "适中": "+0%", "偏快": "+18%"}
    return mapping.get(pace.strip(), "+0%")


def synthesize_edge_tts(text: str, voice: str = "zh-CN-XiaoxiaoNeural", pace: str = "适中") -> bytes:
    """Synthesize speech using free Microsoft Edge Neural TTS and convert to 24kHz mono WAV."""
    if edge_tts is None:
        raise RuntimeError("未安装 edge-tts 库，请先运行 pip install edge-tts。")
    if not text.strip():
        raise ValueError("请输入数字人台词。")
    ffmpeg = audio_manager.find_ffmpeg()
    if not ffmpeg:
        raise RuntimeError("未找到 FFmpeg，无法将 Edge 语音转换为 WAV。")
    voice_code = voice.split(" · ")[0].strip() if " · " in voice else voice.strip()
    if voice_code not in EDGE_VOICE_CODES:
        voice_code = EDGE_VOICE_CODES[0]

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        mp3_path = tmp / "edge.mp3"
        wav_path = tmp / "edge.wav"

        async def _run():
            comm = edge_tts.Communicate(text.strip(), voice_code, rate=_edge_rate(pace))
            await comm.save(str(mp3_path))

        asyncio.run(_run())
        if not mp3_path.is_file() or mp3_path.stat().st_size == 0:
            raise RuntimeError("微软神经语音未返回音频，请检查网络连接。")

        res = subprocess.run(
            [ffmpeg, "-y", "-i", str(mp3_path), "-ac", "1", "-ar", "24000", "-sample_fmt", "s16", str(wav_path)],
            capture_output=True,
            text=True,
            check=False,
        )
        if res.returncode != 0 or not wav_path.is_file():
            raise RuntimeError(f"Edge 语音转码 WAV 失败: {res.stderr.strip()}")
        wav_bytes = wav_path.read_bytes()
        audio_manager.validate_wav_bytes(wav_bytes)
        return wav_bytes


def synthesize_windows_offline(text: str, voice: str = "Microsoft Huihui Desktop", pace: str = "适中") -> bytes:
    """Synthesize speech offline via Windows System.Speech SAPI into WAV."""
    if not text.strip():
        raise ValueError("请输入数字人台词。")
    rate_val = {"偏慢": -2, "适中": 0, "偏快": 2}.get(pace.strip(), 0)
    with tempfile.TemporaryDirectory() as tmpdir:
        wav_path = Path(tmpdir) / "sapi.wav"
        voice_arg = voice.strip().replace("'", "''")
        ps = (
            "Add-Type -AssemblyName System.Speech; "
            "$text = [Console]::In.ReadToEnd().Trim(); "
            "$synth = New-Object System.Speech.Synthesis.SpeechSynthesizer; "
            "try { "
            f"try {{ $synth.SelectVoice('{voice_arg}') }} catch {{}}; "
            f"$synth.Rate = {rate_val}; "
            f"$synth.SetOutputToWaveFile('{wav_path}'); "
            "$synth.Speak($text); "
            "} finally { $synth.Dispose() }"
        )
        _run_powershell(ps, stdin_text=text.strip())
        if not wav_path.is_file() or wav_path.stat().st_size <= 44:
            raise RuntimeError("Windows 离线语音生成失败。")
        wav_bytes = wav_path.read_bytes()
        audio_manager.validate_wav_bytes(wav_bytes)
        return wav_bytes


def is_qwen3_available(
    python_path: Path = QWEN3_PYTHON,
    model_dir: Path = QWEN3_MODEL_DIR,
    worker_script: Path = QWEN3_WORKER,
) -> bool:
    """Return True when the local Qwen3-TTS venv, model weights, and worker exist."""
    return (
        Path(python_path).is_file()
        and (Path(model_dir) / "model.safetensors").is_file()
        and Path(worker_script).is_file()
    )


def default_engine() -> str:
    """Prefer local Qwen3-TTS when installed, otherwise fall back to Edge Neural TTS."""
    return "qwen3" if is_qwen3_available() else "edge"


def _normalize_qwen3_voice(voice: str) -> str:
    code = voice.split(" · ")[0].strip().lower() if voice else ""
    return code if code in QWEN3_VOICE_CODES else QWEN3_VOICE_CODES[0]


def _build_qwen3_instruct(style: str, pace: str) -> str:
    parts = [style.strip()] if style and style.strip() else []
    pace_hint = {
        "偏慢": "语速偏慢，吐字从容清晰",
        "偏快": "语速略快，紧凑明快",
    }.get((pace or "").strip())
    if pace_hint:
        parts.append(pace_hint)
    return "，".join(parts) if parts else "温和、亲切、清晰的课堂讲解"


class _Qwen3WorkerManager:
    """Keeps one local Qwen3-TTS inference process warm across slide generations."""

    def __init__(self) -> None:
        self._req_lock = threading.Lock()
        self._proc_lock = threading.Lock()
        self._proc: subprocess.Popen[str] | None = None
        self._cfg: tuple[str, str, str] | None = None
        self._cuda: bool = True
        self._stdout_queue: queue.Queue[str | None] | None = None
        self._stderr_tail: deque[str] = deque(maxlen=40)

    def _drain_stderr(self, pipe) -> None:
        try:
            for line in pipe:
                text = line.rstrip()
                if text:
                    self._stderr_tail.append(text)
        except Exception:  # noqa: BLE001
            pass

    @staticmethod
    def _pump_stdout(pipe, out_queue: queue.Queue[str | None]) -> None:
        try:
            for line in pipe:
                out_queue.put(line.rstrip())
        except Exception:  # noqa: BLE001
            pass
        finally:
            out_queue.put(None)

    def _kill_proc(self, proc: subprocess.Popen[str] | None) -> None:
        if proc is None:
            return
        try:
            if proc.poll() is None:
                proc.kill()
                proc.wait(timeout=2)
        except Exception:  # noqa: BLE001
            pass
        finally:
            for stream in (proc.stdin, proc.stdout, proc.stderr):
                if stream is not None:
                    try:
                        stream.close()
                    except Exception:  # noqa: BLE001
                        pass

    def _read_line(self, out_queue: queue.Queue[str | None], timeout_sec: float) -> str | None:
        try:
            return out_queue.get(timeout=timeout_sec)
        except queue.Empty:
            return "__TIMEOUT__"

    def _ensure_started(
        self,
        python_path: Path,
        model_dir: Path,
        worker_script: Path,
        *,
        startup_timeout: float = 60.0,
    ) -> tuple[subprocess.Popen[str], queue.Queue[str | None]]:
        cfg = (str(Path(python_path)), str(Path(model_dir)), str(Path(worker_script)))
        with self._proc_lock:
            if (
                self._proc is not None
                and self._proc.poll() is None
                and self._cfg == cfg
                and self._stdout_queue is not None
            ):
                return self._proc, self._stdout_queue

        self.shutdown()
        self._stderr_tail.clear()
        proc = subprocess.Popen(
            [cfg[0], cfg[2], "--serve", "--model-dir", cfg[1]],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            bufsize=1,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        out_q: queue.Queue[str | None] = queue.Queue()
        threading.Thread(target=self._drain_stderr, args=(proc.stderr,), daemon=True).start()
        threading.Thread(target=self._pump_stdout, args=(proc.stdout, out_q), daemon=True).start()

        with self._proc_lock:
            self._proc = proc
            self._cfg = cfg
            self._stdout_queue = out_q

        ready_line = self._read_line(out_q, startup_timeout)
        if ready_line == "__TIMEOUT__":
            self._force_reset(proc)
            tail = "\n".join(self._stderr_tail)
            raise TimeoutError(f"Qwen3-TTS 本地服务启动超时（>{int(startup_timeout)} 秒）: {tail or '显卡/环境响应缓慢'}")
        try:
            ready = json.loads(ready_line) if ready_line else {}
        except json.JSONDecodeError:
            ready = {}
        if not ready.get("ready"):
            self._force_reset(proc)
            tail = "\n".join(self._stderr_tail)
            raise RuntimeError(f"Qwen3-TTS 本地服务启动失败: {ready_line or tail or '未知错误'}")
        self._cuda = bool(ready.get("cuda", True))
        return proc, out_q

    def _force_reset(self, target_proc: subprocess.Popen[str] | None = None) -> None:
        with self._proc_lock:
            proc = self._proc
            if target_proc is not None and proc is not target_proc:
                proc_to_kill = target_proc
            else:
                proc_to_kill = proc
                self._proc = None
                self._cfg = None
                self._stdout_queue = None
        self._kill_proc(proc_to_kill)

    def request(
        self,
        payload: dict,
        *,
        python_path: Path = QWEN3_PYTHON,
        model_dir: Path = QWEN3_MODEL_DIR,
        worker_script: Path = QWEN3_WORKER,
        timeout_sec: float = 120.0,
    ) -> dict:
        with self._req_lock:
            for attempt in range(2):
                proc, out_q = self._ensure_started(python_path, model_dir, worker_script)
                assert proc.stdin is not None
                try:
                    proc.stdin.write(json.dumps(payload, ensure_ascii=True) + "\n")
                    proc.stdin.flush()
                except OSError:
                    self._force_reset(proc)
                    if attempt == 1:
                        tail = "\n".join(self._stderr_tail)
                        raise RuntimeError(f"Qwen3-TTS 推理进程管道断开: {tail or '无输出'}")
                    continue

                effective_timeout = timeout_sec if self._cuda else min(600.0, max(180.0, timeout_sec * 2.5))
                resp_line = self._read_line(out_q, effective_timeout)
                if resp_line == "__TIMEOUT__":
                    self._force_reset(proc)
                    raise TimeoutError(
                        f"Qwen3-TTS 本页生成超过 {int(effective_timeout)} 秒已自动中止，避免软件卡死。"
                        "可缩短台词重试，或改用微软神经语音。"
                    )
                if resp_line:
                    return json.loads(resp_line)
                self._force_reset(proc)
                if attempt == 1:
                    tail = "\n".join(self._stderr_tail)
                    raise RuntimeError(f"Qwen3-TTS 推理进程异常退出: {tail or '无输出'}")
        raise RuntimeError("Qwen3-TTS 请求失败。")

    def shutdown(self) -> None:
        with self._proc_lock:
            proc = self._proc
            self._proc = None
            self._cfg = None
            self._stdout_queue = None
        if proc is None:
            return
        try:
            if proc.poll() is None and proc.stdin is not None:
                proc.stdin.write('{"cmd": "quit"}\n')
                proc.stdin.flush()
                proc.wait(timeout=2)
        except Exception:  # noqa: BLE001
            pass
        self._kill_proc(proc)


_QWEN3_MANAGER = _Qwen3WorkerManager()
atexit.register(_QWEN3_MANAGER.shutdown)


def shutdown_qwen3_worker() -> None:
    """Terminate the persistent Qwen3-TTS background worker if running."""
    _QWEN3_MANAGER.shutdown()


def synthesize_qwen3_local(
    text: str,
    *,
    voice: str = "serena",
    style: str = "温和、亲切、清晰",
    pace: str = "适中",
    python_path: Path = QWEN3_PYTHON,
    model_dir: Path = QWEN3_MODEL_DIR,
    worker_script: Path = QWEN3_WORKER,
    timeout_sec: float | None = None,
) -> bytes:
    """Synthesize Mandarin speech WAV bytes locally via Qwen3-TTS 0.6B CustomVoice."""
    clean = text.strip()
    if not clean:
        raise ValueError("请输入数字人台词。")
    if not is_qwen3_available(python_path, model_dir, worker_script):
        raise FileNotFoundError(
            "尚未安装本地 Qwen3-TTS 高音质模型，请先双击目录内的「一键安装本地Qwen3高音质模型.bat」，或暂时切换微软神经语音。"
        )
    speaker = _normalize_qwen3_voice(voice)
    instruct = _build_qwen3_instruct(style, pace)
    budget = float(timeout_sec) if timeout_sec is not None else max(75.0, min(300.0, len(clean) * 2.5 + 45.0))
    with tempfile.TemporaryDirectory() as tmpdir:
        out_path = Path(tmpdir) / "qwen3_out.wav"
        res = _QWEN3_MANAGER.request(
            {
                "text": clean,
                "speaker": speaker,
                "language": "Chinese",
                "instruct": instruct,
                "out_path": str(out_path),
            },
            python_path=python_path,
            model_dir=model_dir,
            worker_script=worker_script,
            timeout_sec=budget,
        )
        if not res.get("ok") or not out_path.is_file():
            raise RuntimeError(res.get("error") or "Qwen3-TTS 未生成音频文件。")
        wav_bytes = out_path.read_bytes()
        audio_manager.validate_wav_bytes(wav_bytes)
        return wav_bytes


def synthesize_speech(
    text: str,
    *,
    engine: str = "gemini",
    voice: str = "Leda",
    style: str = "温和、亲切、清晰",
    pace: str = "适中",
    model: str = MODELS[0],
    key_path: Path = DEFAULT_KEY_PATH,
    api_key: str | None = None,
) -> bytes:
    """Synthesize Mandarin speech WAV bytes via qwen3, edge, windows, or gemini engine."""
    if engine == "qwen3":
        return synthesize_qwen3_local(text, voice=voice, style=style, pace=pace)
    if engine == "edge":
        return synthesize_edge_tts(text, voice=voice, pace=pace)
    if engine == "windows":
        return synthesize_windows_offline(text, voice=voice, pace=pace)
    if engine == "gemini":
        resolved_key = api_key.strip() if api_key else load_api_key(key_path)
        return gemini_tts.synthesize(
            resolved_key,
            text,
            voice,
            style,
            pace,
            model=model,
        )
    raise ValueError(f"不支持的语音引擎: {engine}")
