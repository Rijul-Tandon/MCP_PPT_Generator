import asyncio
import os
import sys
import json
from dotenv import load_dotenv
import httpx
from groq import AsyncGroq
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

# Load environment variables from .env file
load_dotenv()

async def run_chat():
    groq_api_key = os.getenv("GROQ_API_KEY")
    if not groq_api_key:
        print("Error: GROQ_API_KEY is not set in .env")
        print("Please add it to your .env file at the root of the project.")
        return

    # Initialize Groq client with SSL verification disabled for corporate firewalls
    http_client = httpx.AsyncClient(verify=False)
    llm_client = AsyncGroq(api_key=groq_api_key, http_client=http_client)
    
    # Configure MCP Server parameters for both servers
    excel_params = StdioServerParameters(command=sys.executable, args=["run_excel.py"])
    ppt_params = StdioServerParameters(command=sys.executable, args=["run_ppt.py"])

    print("Starting MCP Servers and connecting to Groq...")
    
    # We use nested context managers to connect to both servers over stdio
    async with stdio_client(excel_params) as (e_read, e_write), \
               stdio_client(ppt_params) as (p_read, p_write):
        async with ClientSession(e_read, e_write) as e_session, \
                   ClientSession(p_read, p_write) as p_session:
            # Initialize connection to both servers
            await e_session.initialize()
            await p_session.initialize()
            print("Connected to Excel and PPT MCP Servers!")
            
            # Get available tools from both MCP servers
            e_tools_response = await e_session.list_tools()
            p_tools_response = await p_session.list_tools()
            
            # Map tools to their respective sessions so we know which server to call
            mcp_tools = []
            tool_sessions = {}
            
            for t in e_tools_response.tools:
                mcp_tools.append(t)
                tool_sessions[t.name] = e_session
                
            for t in p_tools_response.tools:
                mcp_tools.append(t)
                tool_sessions[t.name] = p_session
            
            # Convert MCP tools to Groq's expected JSON schema format
            groq_tools = []
            for t in mcp_tools:
                groq_tools.append({
                    "type": "function",
                    "function": {
                        "name": t.name,
                        "description": t.description,
                        "parameters": t.inputSchema
                    }
                })
            
            print(f"Loaded {len(groq_tools)} tools from the MCP server.")
            print("Type 'exit' or 'quit' to stop.\n")
            print("-" * 50)
            
            # Initialize the conversation history
            messages = [
                {
                    "role": "system",
                    "content": (
                        "You are a top-tier pharmaceutical presentation assistant. "
                        "You have access to MCP tools that can list use cases, read/write context, "
                        "query Excel data with SQL, and modify PowerPoint presentations. "
                        "CRITICAL WORKFLOW INSTRUCTIONS:\n"
                        "1. When asked to build or edit a presentation, you MUST first read the context and use Excel tools to extract CRISP, highly relevant data and insights for charts.\n"
                        "2. You MUST then output a detailed textual plan for the slides. Sub-iterations are expected: you should discuss the plan with the user to make it top notch.\n"
                        "3. You MUST wait for the user to explicitly type 'confirmed' before actually generating or editing the presentation.\n"
                        "4. UNDER NO CIRCUMSTANCES should you execute presentation generation tools without explicit user confirmation of the slide plan first."
                    )
                }
            ]
            
            while True:
                try:
                    user_input = input("\nYou: ")
                    if user_input.strip().lower() in ['exit', 'quit']:
                        break
                    if not user_input.strip():
                        continue
                        
                    messages.append({"role": "user", "content": user_input})
                    
                    # Agent Loop: Keep calling the LLM until it stops asking for tools
                    while True:
                        try:
                            # Call Groq LLM
                            response = await llm_client.chat.completions.create(
                                model="llama-3.3-70b-versatile",
                                messages=messages,
                                tools=groq_tools,
                                tool_choice="auto",
                                parallel_tool_calls=False,
                                max_tokens=4000
                            )
                            
                            response_message = response.choices[0].message
                            
                            # Check if Groq wants to call a tool
                            if response_message.tool_calls:
                                # We must append the assistant's tool-call message to the history
                                messages.append(response_message)
                                
                                for tool_call in response_message.tool_calls:
                                    print(f"\n[AI is running tool: {tool_call.function.name}]")
                                    
                                    # Execute the tool via our local MCP server
                                    args = json.loads(tool_call.function.arguments)
                                    
                                    try:
                                        # Call the tool on the specific session that owns it
                                        target_session = tool_sessions[tool_call.function.name]
                                        result = await target_session.call_tool(tool_call.function.name, arguments=args)
                                        
                                        # Extract text from MCP CallToolResult
                                        tool_output = ""
                                        for content in result.content:
                                            if content.type == "text":
                                                tool_output += content.text
                                        
                                        # Pass the tool's result back to Groq
                                        messages.append({
                                            "role": "tool",
                                            "tool_call_id": tool_call.id,
                                            "name": tool_call.function.name,
                                            "content": tool_output
                                        })
                                        print(f"[Tool finished]")
                                        
                                    except Exception as e:
                                        messages.append({
                                            "role": "tool",
                                            "tool_call_id": tool_call.id,
                                            "name": tool_call.function.name,
                                            "content": f"Error calling tool: {str(e)}"
                                        })
                                        print(f"[Tool failed: {e}]")
                                        
                                # Continue the inner while loop to send tool results back to Groq
                                continue
                                
                            else:
                                # Groq just returned a simple text response without needing tools
                                reply = response_message.content
                                messages.append({"role": "assistant", "content": reply})
                                print(f"\nAssistant: {reply}")
                                break # Exit the agent loop and wait for next user input

                        except Exception as inner_e:
                            error_str = str(inner_e)
                            # Groq specific fallback for LLaMA 3 tool call hallucinations
                            if "failed_generation" in error_str and "<function=" in error_str:
                                import re
                                match = re.search(r'<function=([a-zA-Z0-9_]+)[^\{]*(\{.*?\})</function>', error_str)
                                if match:
                                    func_name = match.group(1)
                                    func_args_str = match.group(2)
                                    print(f"\n[AI is running tool (Fallback Recovery): {func_name}]")
                                    try:
                                        args = json.loads(func_args_str)
                                        target_session = tool_sessions[func_name]
                                        result = await target_session.call_tool(func_name, arguments=args)
                                        tool_output = "".join([c.text for c in result.content if c.type == "text"])
                                        # We spoof a user message containing the tool result so the AI can continue
                                        messages.append({
                                            "role": "user", 
                                            "content": f"The tool '{func_name}' was executed. Here is the result:\n{tool_output}"
                                        })
                                        print(f"[Tool finished]")
                                        continue # Let Groq try again with the fallback result
                                    except Exception as fallback_e:
                                        print(f"[Tool failed during fallback: {fallback_e}]")
                                        break
                            # If it's a different error, raise it to the outer loop
                            raise inner_e

                except KeyboardInterrupt:
                    print("\nExiting...")
                    break
                except Exception as e:
                    print(f"\nAn error occurred: {e}")

if __name__ == "__main__":
    asyncio.run(run_chat())
