import json
from groq import Groq
from src.excel_agent.config import Config
from src.excel_agent.tools import get_registered_tools
from typing import Union, List, Dict
from src.excel_agent.output.abstract_output_handler import AbstractOutputHandler

class LLMInterface:
    def __init__(self, output_handler: AbstractOutputHandler):
        self.client = Groq(api_key=Config.GROQ_API_KEY)
        self.model_name = Config.GROQ_MODEL_NAME
        self.output_handler = output_handler

    def get_tool_call(self, user_query: str, temperature: float = 0.0) -> Union[List[Dict], Dict]:
        if not Config.GROQ_API_KEY:
            self.output_handler.show_error("Groq API key is not configured.")
            return {"error": "Groq API key is not configured."}

        try:
            tools_schema = get_registered_tools()
            if not tools_schema:
                self.output_handler.show_error("No tools registered. Please ensure ExcelHandler methods are decorated with @tool.")
                return {"error": "No tools registered. Please ensure ExcelHandler methods are decorated with @tool."}

            messages = [{"role": "user", "content": user_query}]
            
            chat_completion = self.client.chat.completions.create(
                messages=messages,
                model=self.model_name,
                tools=tools_schema,
                tool_choice="auto",
                temperature=temperature,
            )

            response_dict = chat_completion.to_dict()
            
            if hasattr(chat_completion.choices[0].message, 'tool_calls') and chat_completion.choices[0].message.tool_calls:
                raw_tool_calls = [tc.to_dict() for tc in chat_completion.choices[0].message.tool_calls]
                
                parsed_tool_calls = []
                for tool_call in chat_completion.choices[0].message.tool_calls:
                    try:
                        parsed_tool_calls.append({
                            "tool_name": tool_call.function.name,
                            "tool_parameters": json.loads(tool_call.function.arguments)
                        })
                    except json.JSONDecodeError as e:
                        # Use repr() to safely display the raw arguments string
                        self.output_handler.show_error(f"JSON Parse Error: Failed to parse tool arguments for '{tool_call.function.name}': {e}. Raw arguments: {repr(tool_call.function.arguments)}")
                        return {"error": f"Failed to parse tool arguments: {e}"}
                return parsed_tool_calls
            else:
                self.output_handler.show_warning("No Tool Calls: LLM did not return any tool calls.")
                return {"error": "LLM did not return any tool calls."}

        except Exception as e:
            self.output_handler.show_error(f"API Error: {str(e)}")
            return {"error": f"Error communicating with Groq API: {str(e)}"}

