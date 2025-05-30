# excel_agent/prompts.py

class Prompts:
    """
    Manages prompts for the LLM.
    In tool-calling mode, the primary prompt is constructed directly in agent.py
    and passed as the user message, along with the tool definitions.
    This class is now largely a placeholder or could be used for more complex
    multi-turn conversational prompts if needed in the future.
    """
    # The SYSTEM_PROMPT and FEW_SHOT_EXAMPLES are no longer directly used
    # for code generation. The LLM's behavior is guided by the 'tools' parameter
    # in the API call and the user message constructed in agent.py.

    @staticmethod
    def construct_prompt(user_query: str, file_path: str, sheet_name: str, column_headers: list, all_sheet_names: list) -> str:
        """
        This method is deprecated in tool-calling mode. The context is passed directly
        to the LLM via the user message in agent.py.
        """
        raise NotImplementedError("construct_prompt is deprecated in tool-calling mode. Context is passed directly to LLM.")

