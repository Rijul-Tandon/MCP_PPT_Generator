"""
chat.py - Main MCP Client Entry Point

Connects the LLM to two MCP servers:
  - PPT Server (run_ppt.py)  : PowerPoint read/write/chart operations
  - Excel Server (run_excel.py): Excel data extraction (optional)

Provider is selected via LLM_PROVIDER in .env:
  LLM_PROVIDER=cerebras  (default, higher limits)
  LLM_PROVIDER=groq      (fallback)

Usage:
  python chat.py              # Full mode: PPT + Excel
  python chat.py --no-excel   # Layout mode: PPT only
"""

import argparse
import asyncio
import json
import os
import re
import sys
from dotenv import load_dotenv
import httpx
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

load_dotenv()

# ─────────────────────────────────────────────
# LLM client factory
# Supports Cerebras (primary) and Groq (fallback).
# Both expose an identical async chat.completions.create interface.
# ─────────────────────────────────────────────

def _build_llm_client():
    """
    Construct the async LLM client based on LLM_PROVIDER in .env.
    Returns (client, model_name).
    SSL verification is disabled for corporate firewalls.
    """
    provider = os.getenv("LLM_PROVIDER", "cerebras").lower()

    # Shared SSL-tolerant HTTP client
    http_client = httpx.AsyncClient(verify=False)

    if provider == "cerebras":
        api_key = os.getenv("CEREBRAS_API_KEY")
        if not api_key:
            raise RuntimeError("CEREBRAS_API_KEY not set in .env")
        from cerebras.cloud.sdk import AsyncCerebras
        client = AsyncCerebras(api_key=api_key, http_client=http_client)
        model  = os.getenv("CEREBRAS_MODEL", "gpt-oss-120b")
        return client, model, "cerebras"

    elif provider == "groq":
        api_key = os.getenv("GROQ_API_KEY")
        if not api_key:
            raise RuntimeError("GROQ_API_KEY not set in .env")
        from groq import AsyncGroq
        client = AsyncGroq(api_key=api_key, http_client=http_client)
        model  = os.getenv("GROQ_MODEL", "llama-3.3-70b-versatile")
        return client, model, "groq"

    else:
        raise RuntimeError(f"Unknown LLM_PROVIDER '{provider}'. Use 'cerebras' or 'groq'.")


# ─────────────────────────────────────────────
# System Prompts
# ─────────────────────────────────────────────

SYSTEM_PROMPT_FULL = """You are an expert pharmaceutical analytics presentation assistant.
All MCP communication uses JSON-RPC; XML travels as string values inside JSON.

TOOLS:
  list_use_cases, read_context, list_data_sources
  get_deck_manifest         - compact slide-title index (~300 tokens per deck)
  get_slide_text            - extracts text from a slide (~100 tokens, use instead of getting full XML)
  clone_deck_from_reference - clones a reference deck and keeps only requested slides (100% design fidelity)
  replace_slide_text        - surgically replaces text on a slide while perfectly preserving its styling
  write_slide_chart_xml     - surgically replace only chart XML on a slide

STRICT INTERVIEW WORKFLOW - follow EXACTLY in this order:

STEP 1 - MODE: Ask "New or Refine?". Wait.

STEP 2 - USE CASE: Call list_use_cases. Show options. Ask which. Wait.

STEP 3a - DECK DISCOVERY:
  Call list_data_sources. Call get_deck_manifest for EACH deck found.
  Show: Deck name | Slide count | First 5 titles.
  Ask: "Which deck best matches the visual design you want?" Wait.

STEP 3b - SLIDE SELECTION:
  Show the full title list for the chosen deck.
  Ask: "Which slide indices do you want to clone for the new deck? (e.g. 0, 4, 4, 9)" Wait.

STEP 4 - CONTENT READ:
  For each unique chosen slide index, call get_slide_text to see the existing text blocks.
  Also call read_context to understand the business story.
  Summarise findings in 3-4 sentences.

STEP 5 - SLIDE PLAN (iterate until confirmed):
  Draft a detailed plan for the NEW deck. For each slide:
    - Target Slide Index (0-based in the new deck)
    - Cloned from reference slide index X
    - Proposed Text Replacements (Old Text -> New Text)
  Revise until user types exactly: confirmed

STEP 6 - GENERATE (only after confirmed):
  a) Call clone_deck_from_reference to create the output .pptx file, passing the exact slide_mapping array.
  b) For each slide in the NEW deck that needs text changes:
     - Call replace_slide_text on the NEW output file and NEW slide index.
       Pass a dictionary mapping {"Exact Old Text": "New Text"}.
  c) For chart-only updates: author <c:chartSpace> XML and call write_slide_chart_xml.
  Report the output file path.

GOLDEN RULES:
- NEVER call any generation tool before "confirmed".
- ONE slide per get_slide_text call - never batch.
- Speak like a senior consultant.
"""

SYSTEM_PROMPT_LAYOUT_ONLY = """You are an expert pharmaceutical analytics presentation assistant.
Layout-Only Mode - Excel server NOT connected. Build a branded deck shell using a reference deck.

TOOLS:
  list_use_cases, read_context, list_data_sources
  get_deck_manifest, get_slide_text
  clone_deck_from_reference, replace_slide_text

STRICT INTERVIEW WORKFLOW:

STEP 1 - MODE: Ask "New or Refine?". Wait.
STEP 2 - USE CASE: Call list_use_cases. Ask which. Wait.
STEP 3a - DECK DISCOVERY:
  Call list_data_sources. Call get_deck_manifest for each deck.
  Show: name | count | first 5 titles. Ask which style deck to clone. Wait.
STEP 3b - SLIDE SELECTION:
  Show full title list. Ask which indices to clone for the new deck (e.g., 0, 4, 4, 9). Wait.
STEP 4 - CONTENT READ:
  Call get_slide_text for each chosen slide.
  Study the extracted text strings to understand the layout and placeholders.
  Summarise in 3-4 sentences.
STEP 5 - SLIDE PLAN (iterate):
  Draft plan. For each slide in the NEW deck:
    - Target Slide Index (0-based)
    - Cloned from reference index X
    - Text Replacements (Old Text -> New Text)
  Revise until user types: confirmed
STEP 6 - GENERATE (only after confirmed):
  Call clone_deck_from_reference to create the new PPTX file perfectly matching the design.
  For each new slide that needs updated text:
    - Call replace_slide_text on the NEW output file and NEW slide index.
  Report the output path.

GOLDEN RULES:
- No Excel tools - not connected.
- NEVER call generation tools before "confirmed".
- ONE slide per get_slide_text call.
"""


# ─────────────────────────────────────────────────────────────────────────────
# Robust JSON parser for LLM-authored payloads containing XML strings.
#
# LLaMA / Cerebras models sometimes embed OOXML in function call arguments
# without proper JSON escaping. The parser tries progressively more
# aggressive repair strategies before giving up.
# ─────────────────────────────────────────────────────────────────────────────

def _try_parse_json(raw: str) -> dict | None:
    """Try multiple repair strategies to parse potentially malformed JSON."""
    # Strategy 1: Direct (LLM was well-behaved)
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass

    # Strategy 2: Fix raw newlines / carriage returns / tabs inside strings
    try:
        fixed = re.sub(r'(?<!\\)\n', r'\\n', raw)
        fixed = re.sub(r'(?<!\\)\r', r'\\r', fixed)
        fixed = re.sub(r'(?<!\\)\t', r'\\t', fixed)
        return json.loads(fixed)
    except json.JSONDecodeError:
        pass

    # Strategy 3: Fix lone backslashes not part of a valid escape sequence
    try:
        fixed = re.sub(r'\\(?!["\\/bfnrtu])', r'\\\\', raw)
        return json.loads(fixed)
    except json.JSONDecodeError:
        pass

    # Strategy 4: Combine strategy 2 + 3
    try:
        fixed = re.sub(r'(?<!\\)\n', r'\\n', raw)
        fixed = re.sub(r'(?<!\\)\r', r'\\r', fixed)
        fixed = re.sub(r'\\(?!["\\/bfnrtu])', r'\\\\', fixed)
        return json.loads(fixed)
    except json.JSONDecodeError:
        pass

    return None


def _parse_llama_xml_tool_call(error_str: str) -> tuple[str | None, dict | None]:
    """
    Extract (func_name, args_dict) from a LLaMA XML-style function call
    embedded in a Groq/Cerebras error string.

    Uses greedy match + re.DOTALL so multi-line XML inside the JSON blob
    is captured correctly. Falls back to _try_parse_json for repair.
    """
    match = re.search(
        r"<function=([a-zA-Z0-9_]+)\s*(\{.*\})\s*</function>",
        error_str,
        re.DOTALL,
    )
    if not match:
        return None, None

    func_name = match.group(1)
    args = _try_parse_json(match.group(2))
    return func_name, args  # args may be None if all repair strategies fail


# ─────────────────────────────────────────────────────────────────────────────
# LLM completion wrapper
# Normalises the interface between Cerebras and Groq since their parameter
# names differ slightly (max_completion_tokens vs max_tokens).
# ─────────────────────────────────────────────────────────────────────────────

async def _complete(client, provider: str, model: str, messages: list, tools: list) -> object:
    """Call chat.completions.create with provider-appropriate parameters."""
    common = dict(
        model=model,
        messages=messages,
        tools=tools,
        tool_choice="auto",
        stream=False,
    )

    if provider == "cerebras":
        return await client.chat.completions.create(
            **common,
            max_completion_tokens=8192,
        )
    else:  # groq
        return await client.chat.completions.create(
            **common,
            parallel_tool_calls=False,
            max_tokens=4096,
        )


# ─────────────────────────────────────────────────────────────────────────────
# Main entry point
# ─────────────────────────────────────────────────────────────────────────────

async def run_chat(use_excel: bool) -> None:
    """Boot MCP servers and start the interactive chat loop."""

    try:
        llm_client, model, provider = _build_llm_client()
    except RuntimeError as e:
        print(f"Error: {e}")
        return

    ppt_params = StdioServerParameters(command=sys.executable, args=["run_ppt.py"])

    mode_label   = "Full Mode (PPT + Excel)" if use_excel else "Layout-Only Mode (PPT only)"
    system_prompt = SYSTEM_PROMPT_FULL if use_excel else SYSTEM_PROMPT_LAYOUT_ONLY

    print(f"\nStarting MCP Servers... [{mode_label}]")
    print(f"LLM Provider : {provider.upper()} — {model}")

    async def _build_registry(p_session, e_session):
        await p_session.initialize()
        groq_tools, tool_sessions = [], {}
        for t in (await p_session.list_tools()).tools:
            tool_sessions[t.name] = p_session
            groq_tools.append({"type": "function", "function": {
                "name": t.name, "description": t.description, "parameters": t.inputSchema
            }})
        if e_session:
            await e_session.initialize()
            for t in (await e_session.list_tools()).tools:
                tool_sessions[t.name] = e_session
                groq_tools.append({"type": "function", "function": {
                    "name": t.name, "description": t.description, "parameters": t.inputSchema
                }})
        return groq_tools, tool_sessions

    async def _chat_loop(groq_tools, tool_sessions):
        print(f"Connected! Loaded {len(groq_tools)} tools.\nType 'exit' or 'quit' to stop.\n" + "-" * 50)
        messages = [{"role": "system", "content": system_prompt}]
        messages.append({"role": "user", "content": "Hello, let's get started."})
        await _agent_turn(llm_client, provider, model, groq_tools, tool_sessions, messages)

        while True:
            try:
                user_input = input("\nYou: ").strip()
                if user_input.lower() in ("exit", "quit"):
                    print("Goodbye!")
                    break
                if not user_input:
                    continue
                messages.append({"role": "user", "content": user_input})
                await _agent_turn(llm_client, provider, model, groq_tools, tool_sessions, messages)
            except KeyboardInterrupt:
                print("\nExiting...")
                break
            except Exception as e:
                print(f"\nAn error occurred: {e}")

    async with stdio_client(ppt_params) as (p_read, p_write):
        async with ClientSession(p_read, p_write) as p_session:
            if use_excel:
                excel_params = StdioServerParameters(command=sys.executable, args=["run_excel.py"])
                async with stdio_client(excel_params) as (e_read, e_write):
                    async with ClientSession(e_read, e_write) as e_session:
                        tools, sessions = await _build_registry(p_session, e_session)
                        await _chat_loop(tools, sessions)
            else:
                tools, sessions = await _build_registry(p_session, None)
                await _chat_loop(tools, sessions)


# ─────────────────────────────────────────────────────────────────────────────
# Agent turn — handles tool calls and LLaMA XML-style fallback
# ─────────────────────────────────────────────────────────────────────────────

async def _agent_turn(llm_client, provider: str, model: str, groq_tools, tool_sessions, messages) -> None:
    """
    Inner loop: LLM -> tool calls -> results -> repeat until plain text reply.

    Handles two known quirks of LLaMA-family models on Groq/Cerebras:
      1. XML-style function calls (<function=name {...}></function>)
      2. OOXML embedded in JSON strings without proper escaping
    """
    while True:
        try:
            response = await _complete(llm_client, provider, model, messages, groq_tools)
            msg = response.choices[0].message

            # ── Proper JSON tool call ─────────────────────────────────────
            if msg.tool_calls:
                messages.append(msg)
                for tc in msg.tool_calls:
                    print(f"\n[🔧 Tool: {tc.function.name}]")
                    args = json.loads(tc.function.arguments)
                    try:
                        result = await tool_sessions[tc.function.name].call_tool(
                            tc.function.name, arguments=args
                        )
                        output = "".join(c.text for c in result.content if c.type == "text")
                        print(f"[✅ Done: {tc.function.name}]")
                    except Exception as e:
                        output = f"Tool error: {e}"
                        print(f"[❌ Failed: {tc.function.name} — {e}]")

                    messages.append({
                        "role": "tool",
                        "tool_call_id": tc.id,
                        "name": tc.function.name,
                        "content": output,
                    })
                continue

            # ── Plain text reply ──────────────────────────────────────────
            reply = msg.content or ""
            messages.append({"role": "assistant", "content": reply})
            print(f"\nAssistant: {reply}")
            break

        except Exception as e:
            error_str = str(e)

            # ── Fallback: LLaMA XML-style tool call ──────────────────────
            if "failed_generation" in error_str and "<function=" in error_str:
                func_name, args = _parse_llama_xml_tool_call(error_str)

                if func_name and args is not None:
                    print(f"\n[🔧 Tool (fallback): {func_name}]")
                    try:
                        result = await tool_sessions[func_name].call_tool(func_name, arguments=args)
                        output = "".join(c.text for c in result.content if c.type == "text")
                        messages.append({"role": "user", "content": f"Tool '{func_name}' result:\n{output}"})
                        print(f"[✅ Done (fallback): {func_name}]")
                        continue
                    except Exception as fe:
                        print(f"[❌ Fallback tool call failed: {fe}]")
                        messages.append({"role": "user", "content": (
                            f"Tool '{func_name}' failed: {fe}. "
                            "Retry with a proper JSON tool call. "
                            "When embedding XML inside JSON strings, escape every "
                            'double-quote as \\" and every newline as \\n.'
                        )})
                        continue

                else:
                    # Name found but args unparseable — guide the LLM
                    guidance = (
                        f"Your call to '{func_name}' had malformed JSON arguments. "
                        if func_name else
                        "Your response contained a malformed function call. "
                    )
                    messages.append({"role": "user", "content": (
                        guidance +
                        "Please retry using a standard JSON tool call with all "
                        "XML special characters properly escaped."
                    )})
                    continue

            raise  # Unknown error — bubble up


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Pharma Presentation Agent — MCP Client")
    parser.add_argument(
        "--no-excel", action="store_true",
        help="Layout-Only mode: skip the Excel server."
    )
    args = parser.parse_args()
    asyncio.run(run_chat(use_excel=not args.no_excel))
