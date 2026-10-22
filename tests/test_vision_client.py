import pytest
from unittest.mock import patch, MagicMock
import httpx
from bom_enricher.vision_client import VisionLLMClient


def test_vision_client_is_available_true():
    client = VisionLLMClient(base_url="http://localhost:11434")
    with patch("httpx.get") as mock_get:
        mock_get.return_value = MagicMock(status_code=200)
        assert client.is_available() is True
        mock_get.assert_called_once_with("http://localhost:11434/api/tags", timeout=1.5)


def test_vision_client_is_available_false_on_connection_error():
    client = VisionLLMClient(base_url="http://localhost:11434")
    with patch("httpx.get", side_effect=httpx.ConnectError("Connection refused")):
        assert client.is_available() is False


def test_vision_client_is_available_false_on_non_200():
    client = VisionLLMClient(base_url="http://localhost:11434")
    with patch("httpx.get") as mock_get:
        mock_get.return_value = MagicMock(status_code=500)
        assert client.is_available() is False


def test_vision_client_extract_specs_success():
    client = VisionLLMClient(base_url="http://localhost:11434")
    mock_response = MagicMock(status_code=200)
    mock_response.json.return_value = {
        "message": {
            "content": '{"material_type": "CON", "confidence": 0.95, "manufacturer": "Degson", "specs": {"pitch": "8.2mm", "pin_number": "12 pin", "type_connector": "Wire to Wire Klemens", "mounting": "Cable", "pin_type": "Dişi"}}'
        }
    }
    with patch("httpx.post", return_value=mock_response) as mock_post:
        result = client.extract_specs_from_images(
            images=["base64_dummy_1"],
            target_pn="DG8HS-1.2-12P-17-100SH/AH"
        )
        assert result is not None
        assert result["material_type"] == "CON"
        assert result["confidence"] == 0.95
        assert result["manufacturer"] == "Degson"
        assert result["specs"]["pitch"] == "8.2mm"
        assert result["specs"]["pin_number"] == "12 pin"
        mock_post.assert_called_once()
        args, kwargs = mock_post.call_args
        assert args[0] == "http://localhost:11434/api/chat"
        payload = kwargs["json"]
        assert payload["model"] in ("llama3.2-vision:3b", "moondream:latest")
        assert payload["format"] == "json"
        assert payload["stream"] is False
        assert kwargs["timeout"] == 60.0


def test_vision_client_extract_specs_graceful_timeout():
    client = VisionLLMClient(base_url="http://localhost:11434", timeout=1.0)
    with patch("httpx.post", side_effect=httpx.TimeoutException("Timeout")):
        result = client.extract_specs_from_images(images=["dummy"], target_pn="TEST-PN")
        assert result is None  # Must not crash, return None


def test_vision_client_extract_specs_http_error():
    client = VisionLLMClient(base_url="http://localhost:11434")
    mock_response = MagicMock(status_code=500)
    mock_response.raise_for_status.side_effect = httpx.HTTPStatusError("Server error", request=MagicMock(), response=mock_response)
    with patch("httpx.post", return_value=mock_response):
        result = client.extract_specs_from_images(images=["dummy"], target_pn="TEST-PN")
        assert result is None


def test_vision_client_extract_specs_markdown_wrapped_json():
    client = VisionLLMClient(base_url="http://localhost:11434")
    mock_response = MagicMock(status_code=200)
    mock_response.json.return_value = {
        "message": {
            "content": '```json\n{"material_type": "RES", "confidence": 0.88, "manufacturer": "Yageo", "specs": {"resistance": "10k", "tolerance": "1%"}}\n```'
        }
    }
    with patch("httpx.post", return_value=mock_response):
        result = client.extract_specs_from_images(images=["dummy"], target_pn="RC0805FR-0710KL")
        assert result is not None
        assert result["material_type"] == "RES"
        assert result["specs"]["resistance"] == "10k"


def test_vision_client_extract_specs_invalid_json():
    client = VisionLLMClient(base_url="http://localhost:11434")
    mock_response = MagicMock(status_code=200)
    mock_response.json.return_value = {
        "message": {
            "content": "Not a valid json response at all"
        }
    }
    with patch("httpx.post", return_value=mock_response):
        result = client.extract_specs_from_images(images=["dummy"], target_pn="TEST-PN")
        assert result is None


def test_vision_client_extract_specs_missing_required_keys():
    client = VisionLLMClient(base_url="http://localhost:11434")
    mock_response = MagicMock(status_code=200)
    mock_response.json.return_value = {
        "message": {
            "content": '{"unexpected_key": "some_value"}'
        }
    }
    with patch("httpx.post", return_value=mock_response):
        result = client.extract_specs_from_images(images=["dummy"], target_pn="TEST-PN")
        assert result is None


def test_vision_client_extract_specs_empty_images_or_pn():
    client = VisionLLMClient(base_url="http://localhost:11434")
    with patch("httpx.post") as mock_post:
        assert client.extract_specs_from_images(images=[], target_pn="TEST-PN") is None
        assert client.extract_specs_from_images(images=["dummy"], target_pn="") is None
        assert client.extract_specs_from_images(images=["dummy"], target_pn="   ") is None
        mock_post.assert_not_called()


def test_vision_client_prompt_includes_target_pn_and_hint():
    client = VisionLLMClient(base_url="http://localhost:11434", model="qwen2-vl:2b")
    mock_response = MagicMock(status_code=200)
    mock_response.json.return_value = {
        "message": {
            "content": '{"material_type": "CON", "confidence": 0.9, "manufacturer": "TE", "specs": {"pin_number": "8"}}'
        }
    }
    with patch("httpx.post", return_value=mock_response) as mock_post:
        result = client.extract_specs_from_images(
            images=["img1", "img2"],
            target_pn="XYZ-1234",
            material_type_hint="CON"
        )
        assert result is not None
        assert result["specs"]["pin_number"] == "8"

        call_json = mock_post.call_args[1]["json"]
        assert call_json["model"] == "qwen2-vl:2b"
        messages = call_json["messages"]
        all_content = " ".join(m.get("content", "") for m in messages)
        assert "XYZ-1234" in all_content
        assert "CON" in all_content
        # Check that images were passed
        images_passed = []
        for m in messages:
            if "images" in m:
                images_passed.extend(m["images"])
        assert images_passed == ["img1", "img2"]


def test_vision_client_custom_init_parameters():
    client = VisionLLMClient(
        base_url="http://localhost:11434/",
        model="custom-vlm",
        timeout=30.0,
    )
    assert client.base_url == "http://localhost:11434"
    assert client.model == "custom-vlm"
    assert client.timeout == 30.0


def test_vision_client_json_repair_unterminated_string():
    client = VisionLLMClient(base_url="http://localhost:11434")
    # Truncated string simulating VLM stopping mid-generation
    broken_content = '{"material_type": "CON", "confidence": 0.85, "specs": {"pitch": "2.54mm", "pin_number": "4 pin", "type_connector'
    mock_resp = MagicMock(status_code=200)
    mock_resp.json.return_value = {"message": {"content": broken_content}}
    with patch("httpx.post", return_value=mock_resp):
        res = client.extract_specs_from_images(images=["dummy"], target_pn="HDR-04")
        assert res is not None
        assert res["material_type"] == "CON"
        assert res["specs"]["pitch"] == "2.54mm"
        assert res["specs"]["pin_number"] == "4 pin"


def test_vision_client_placeholder_suppression():
    client = VisionLLMClient(base_url="http://localhost:11434")
    # VLM copying placeholder dots from prompt
    raw_content = '{"material_type": "CON", "manufacturer": "...", "specs": {"pitch": "...", "pin_number": "12 pin", "mounting": "<TH or SMD>"}}'
    mock_resp = MagicMock(status_code=200)
    mock_resp.json.return_value = {"message": {"content": raw_content}}
    with patch("httpx.post", return_value=mock_resp):
        res = client.extract_specs_from_images(images=["dummy"], target_pn="CONN-12")
        assert res is not None
        assert res["specs"]["pin_number"] == "12 pin"
        assert "pitch" not in res["specs"]
        assert "mounting" not in res["specs"]
        assert res["manufacturer"] is None


def test_vision_client_heuristic_text_fallback():
    client = VisionLLMClient(base_url="http://localhost:11434")
    # VLM returned plain text without JSON
    raw_content = "The engineering drawing displays a Connector header with pitch: 1.27mm, 6 pins, through-hole mounting."
    mock_resp = MagicMock(status_code=200)
    mock_resp.json.return_value = {"message": {"content": raw_content}}
    with patch("httpx.post", return_value=mock_resp):
        res = client.extract_specs_from_images(images=["dummy"], target_pn="CONN-06")
        assert res is not None
        assert res["material_type"] == "CON"
        assert res["specs"]["pitch"] == "1.27mm"
        assert res["specs"]["pin_number"] == "6 pin"
        assert res["specs"]["mounting"] == "TH"


def test_vision_llm_client_status_tracking():
    from src.bom_enricher.vision_llm import VisionLLMClient
    client = VisionLLMClient()
    assert client.last_status == "uninitialized"


def test_extract_specs_from_text_timeout_handling(monkeypatch):
    import httpx
    from src.bom_enricher.vision_llm import VisionLLMClient
    client = VisionLLMClient()

    def mock_post(*args, **kwargs):
        raise httpx.TimeoutException("Read timed out")

    monkeypatch.setattr(httpx, "post", mock_post)
    res = client.extract_specs_from_text("WR-PHD SMT Socket Header", timeout=1.0)
    assert res is None
    assert client.last_status == "timeout"


def test_extract_specs_from_text_connection_error_handling(monkeypatch):
    import httpx
    from src.bom_enricher.vision_llm import VisionLLMClient
    client = VisionLLMClient()

    def mock_post(*args, **kwargs):
        raise httpx.ConnectError("Connection refused")

    monkeypatch.setattr(httpx, "post", mock_post)
    res = client.extract_specs_from_text("WR-PHD SMT Socket Header", timeout=1.0)
    assert res is None
    assert client.last_status == "connection_error"


def test_extract_specs_from_text_empty_response(monkeypatch):
    from unittest.mock import MagicMock
    import httpx
    from src.bom_enricher.vision_llm import VisionLLMClient
    client = VisionLLMClient()

    mock_resp = MagicMock(status_code=200)
    mock_resp.json.return_value = {"message": {"content": ""}}
    monkeypatch.setattr(httpx, "post", lambda *args, **kwargs: mock_resp)

    res = client.extract_specs_from_text("WR-PHD SMT Socket Header", timeout=1.0)
    assert res is None
    assert client.last_status == "empty_response"


def test_extract_specs_from_text_parse_error(monkeypatch):
    from unittest.mock import MagicMock
    import httpx
    from src.bom_enricher.vision_llm import VisionLLMClient
    client = VisionLLMClient()

    mock_resp = MagicMock(status_code=200)
    mock_resp.json.return_value = {"message": {"content": "This is completely unparseable gibberish with no json"}}
    monkeypatch.setattr(httpx, "post", lambda *args, **kwargs: mock_resp)

    res = client.extract_specs_from_text("WR-PHD SMT Socket Header", timeout=1.0)
    assert res is None
    assert client.last_status == "parse_error"


def test_extract_specs_from_text_success_status(monkeypatch):
    from unittest.mock import MagicMock
    import httpx
    from src.bom_enricher.vision_llm import VisionLLMClient
    client = VisionLLMClient()

    mock_resp = MagicMock(status_code=200)
    mock_resp.json.return_value = {
        "message": {
            "content": '{"material_type": "CON", "confidence": 0.95, "specs": {"pin_number": "4", "pin_type": "Dişi"}}'
        }
    }
    monkeypatch.setattr(httpx, "post", lambda *args, **kwargs: mock_resp)

    res = client.extract_specs_from_text("WR-PHD SMT Socket Header", timeout=1.0)
    assert res is not None
    assert client.last_status == "success"
    assert res["material_type"] == "CON"
    assert res["specs"]["pin_type"] == "Dişi"


def test_extract_specs_from_images_status_tracking():
    from unittest.mock import MagicMock
    import httpx
    from src.bom_enricher.vision_llm import VisionLLMClient
    client = VisionLLMClient()

    with patch("httpx.post", side_effect=httpx.TimeoutException("Timeout")):
        res = client.extract_specs_from_images(images=["dummy"], target_pn="PN-123")
        assert res is None
        assert client.last_status == "timeout"

    with patch("httpx.post", side_effect=httpx.ConnectError("Refused")):
        res = client.extract_specs_from_images(images=["dummy"], target_pn="PN-123")
        assert res is None
        assert client.last_status == "connection_error"



