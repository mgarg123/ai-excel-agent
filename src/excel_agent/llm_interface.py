import json
from groq import Groq
from src.excel_agent.config import Config # MODIFIED
from src.excel_agent.tools import get_registered_tools # MODIFIED
from typing import Union, List, Dict
# from rich.console import Console # REMOVED
from src.excel_agent.output.abstract_output_handler import AbstractOutputHandler # MODIFIED: NEW IMPORT

class LLMInterface:
    def __init__(self, output_handler: AbstractOutputHandler): # MODIFIED
        self.client = Groq(api_key=Config.GROQ_API_KEY)
        self.model_name = Config.GROQ_MODEL_NAME
        self.output_handler = output_handler # NEW: Store the output handler

    def get_tool_call(self, user_query: str, temperature: float = 0.0) -> Union[List[Dict], Dict]:
        if not Config.GROQ_API_KEY:
            self.output_handler.show_error("Groq API key is not configured.") # MODIFIED
            return {"error": "Groq API key is not configured."}

        try:
            tools_schema = get_registered_tools()
            if not tools_schema:
                self.output_handler.show_error("No tools registered. Please ensure ExcelHandler methods are decorated with @tool.") # MODIFIED
                return {"error": "No tools registered. Please ensure ExcelHandler methods are decorated with @tool."}

            # self.output_handler.print_message(f"Sending to Groq API:", style='warning') # MODIFIED
            # self.output_handler.print_message(f"Model: {self.model_name}", style='warning') # MODIFIED
            # self.output_handler.print_message(f"Query: {user_query}", style='warning') # MODIFIED
            # self.output_handler.print_message(f"Tool Schema being sent: {json.dumps(tools_schema, indent=2)}", style='warning') # MODIFIED
            
            messages = [{"role": "user", "content": user_query}]
            # self.output_handler.print_message(f"Messages: {json.dumps(messages, indent=2)}", style='warning') # MODIFIED
            
            chat_completion = self.client.chat.completions.create(
                messages=messages,
                model=self.model_name,
                tools=tools_schema,
                tool_choice="auto",
                temperature=temperature,
            )

            response_dict = chat_completion.to_dict()
            # self.output_handler.print_message(f"Raw API Response: {json.dumps(response_dict, indent=2)}", style='warning') # MODIFIED
            
            if hasattr(chat_completion.choices[0].message, 'tool_calls') and chat_completion.choices[0].message.tool_calls:
                raw_tool_calls = [tc.to_dict() for tc in chat_completion.choices[0].message.tool_calls]
                # self.output_handler.print_message(f"Raw Tool Calls: {json.dumps(raw_tool_calls, indent=2)}", style='warning') # MODIFIED
                
                parsed_tool_calls = []
                for tool_call in chat_completion.choices[0].message.tool_calls:
                    try:
                        parsed_tool_calls.append({
                            "tool_name": tool_call.function.name,
                            "tool_parameters": json.loads(tool_call.function.arguments)
                        })
                    except json.JSONDecodeError as e:
                        self.output_handler.show_error(f"JSON Parse Error: Failed to parse tool arguments for '{tool_call.function.name}': {e}. Raw arguments: {tool_call.function.arguments}") # MODIFIED
                        return {"error": f"Failed to parse tool arguments: {e}"}
                # self.output_handler.print_message(f"Parsed Tool Calls: {json.dumps(parsed_tool_calls, indent=2)}", style='warning') # MODIFIED
                return parsed_tool_calls
            else:
                self.output_handler.show_warning("No Tool Calls: LLM did not return any tool calls.") # MODIFIED
                return {"error": "LLM did not return any tool calls."}

        except Exception as e:
            self.output_handler.show_error(f"API Error: {str(e)}") # MODIFIED
            return {"error": f"Error communicating with Groq API: {str(e)}"}
