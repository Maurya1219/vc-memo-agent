from typing import TypedDict

from langchain_community.document_loaders import PyPDFLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_openai import OpenAIEmbeddings, ChatOpenAI
from langchain_community.vectorstores import FAISS
from langchain_community.tools.tavily_search import TavilySearchResults
from langchain_classic.agents import create_openai_functions_agent, AgentExecutor
from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
from langchain_core.tools import Tool
from langchain_core.documents import Document
from langgraph.graph import StateGraph, END
import pandas as pd
import os

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
TAVILY_API_KEY = os.getenv("TAVILY_API_KEY")

embeddings = OpenAIEmbeddings(openai_api_key=OPENAI_API_KEY)
llm = ChatOpenAI(model="gpt-4o", temperature=0, openai_api_key=OPENAI_API_KEY)

vectorstore = None

def ingest_pdf(path: str) -> int:
    global vectorstore
    loader = PyPDFLoader(path)
    docs = loader.load()
    splitter = RecursiveCharacterTextSplitter(chunk_size=1000, chunk_overlap=150)
    chunks = splitter.split_documents(docs)
    if vectorstore is None:
        vectorstore = FAISS.from_documents(chunks, embeddings)
    else:
        vectorstore.add_documents(chunks)
    return len(chunks)

def ingest_excel(path: str) -> int:
    global vectorstore
    xl = pd.ExcelFile(path)
    all_chunks = []
    splitter = RecursiveCharacterTextSplitter(chunk_size=1000, chunk_overlap=150)
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        text = f"Sheet: {sheet}\n{df.to_string(index=False)}"
        doc = Document(page_content=text, metadata={"source": path, "sheet": sheet})
        chunks = splitter.split_documents([doc])
        all_chunks.extend(chunks)
    if vectorstore is None:
        vectorstore = FAISS.from_documents(all_chunks, embeddings)
    else:
        vectorstore.add_documents(all_chunks)
    return len(all_chunks)

def make_tools():
    retriever = vectorstore.as_retriever(search_kwargs={"k": 5})
    def search_docs(q: str) -> str:
        docs = retriever.invoke(q)
        return "\n\n".join([d.page_content for d in docs])
    tavily = TavilySearchResults(
        max_results=3,
        name="web_search",
        description="Search the web for market data, competitors, and industry trends.",
        **({"tavily_api_key": TAVILY_API_KEY} if TAVILY_API_KEY else {}),
    )
    return [
        Tool(name="search_documents", func=search_docs,
             description="Search uploaded pitch decks, financials, and market docs."),
        tavily,
    ]

def run_agent(system_prompt: str, question: str) -> str:
    tools = make_tools()
    prompt = ChatPromptTemplate.from_messages([
        ("system", system_prompt),
        ("human", "{input}"),
        MessagesPlaceholder(variable_name="agent_scratchpad")
    ])
    agent = create_openai_functions_agent(llm=llm, tools=tools, prompt=prompt)
    executor = AgentExecutor(agent=agent, tools=tools, verbose=False, max_iterations=4)
    result = executor.invoke({"input": question})
    return result["output"]

def market_agent() -> str:
    return run_agent(
        system_prompt="""You are a market research analyst at a VC firm.
Your job: analyze the market opportunity for this company.
Cover: TAM/SAM/SOM, market growth rate, key trends driving the market,
regulatory environment, and timing — why is now the right time?
Be specific with numbers. Search the web for current market data.
Format as a clear section titled 'Market Analysis'.""",
        question="Analyze the market opportunity, TAM/SAM/SOM, growth trends, and market timing for this company."
    )

def founder_agent() -> str:
    return run_agent(
        system_prompt="""You are a talent and team analyst at a VC firm.
Your job: evaluate the founding team and key personnel.
Cover: founder backgrounds, relevant domain expertise, prior exits or notable experience,
team completeness (any critical gaps?), and advisors.
Search the web for any public information on the founders.
Format as a clear section titled 'Team Assessment'.""",
        question="Evaluate the founding team, their backgrounds, domain expertise, and any gaps in the team."
    )

def finance_agent() -> str:
    return run_agent(
        system_prompt="""You are a financial analyst at a VC firm.
Your job: analyze the financial profile of this company.
Cover: current revenue, MRR/ARR, growth rate, burn rate, runway, unit economics
(CAC, LTV, LTV/CAC ratio), gross margins, and fundraising history.
If numbers are missing, flag them explicitly — a VC needs to know what's not there.
Format as a clear section titled 'Financial Analysis'.""",
        question="Analyze the company's financials including revenue, growth, burn rate, runway, and unit economics."
    )

def risk_agent() -> str:
    return run_agent(
        system_prompt="""You are a risk analyst at a VC firm.
Your job: identify the key risks and red flags for this investment.
Cover: market risks, competitive risks, execution risks, team risks,
technology risks, regulatory risks, and any red flags in the documents.
Be honest and direct — a VC needs an unfiltered view.
Format as a clear section titled 'Risk Assessment'.""",
        question="What are the key risks, red flags, and concerns for this investment?"
    )

class MemoState(TypedDict, total=False):
    market: str
    founder: str
    finance: str
    risk: str
    memo: str

def _node_market(_: MemoState) -> dict:
    print("Running market agent...")
    return {"market": market_agent()}

def _node_founder(_: MemoState) -> dict:
    print("Running founder agent...")
    return {"founder": founder_agent()}

def _node_finance(_: MemoState) -> dict:
    print("Running finance agent...")
    return {"finance": finance_agent()}

def _node_risk(_: MemoState) -> dict:
    print("Running risk agent...")
    return {"risk": risk_agent()}

def _node_writer(state: MemoState) -> dict:
    print("Running writer agent...")
    market = state["market"]
    founder = state["founder"]
    finance = state["finance"]
    risk = state["risk"]
    writer_prompt = f"""You are a senior VC partner writing a formal investment memo.
You have received analysis from four specialist analysts. Synthesize their findings into
a single, well-structured investment memo a partner could present at a Monday meeting.

Use this exact structure:
# Investment Memo

## Executive Summary
(2-3 sentences: what the company does, stage, ask, and your headline recommendation)

## Market Analysis
{market}

## Team Assessment
{founder}

## Financial Analysis
{finance}

## Risk Assessment
{risk}

## Investment Recommendation
(Clear pass/invest/watch recommendation with 3 bullet point rationale)

Write in clear, direct VC memo style. No fluff. Be specific."""
    memo = llm.invoke(writer_prompt).content
    return {"memo": memo}

_memo_builder = StateGraph(MemoState)
_memo_builder.add_node("market", _node_market)
_memo_builder.add_node("founder", _node_founder)
_memo_builder.add_node("finance", _node_finance)
_memo_builder.add_node("risk", _node_risk)
_memo_builder.add_node("writer", _node_writer)
_memo_builder.set_entry_point("market")
_memo_builder.add_edge("market", "founder")
_memo_builder.add_edge("founder", "finance")
_memo_builder.add_edge("finance", "risk")
_memo_builder.add_edge("risk", "writer")
_memo_builder.add_edge("writer", END)
memo_graph = _memo_builder.compile()

def ask_question(question: str) -> dict:
    if vectorstore is None:
        return {"answer": "No documents uploaded yet.", "sources": []}
    tools = make_tools()
    prompt = ChatPromptTemplate.from_messages([
        ("system", """You are an expert VC analyst. Answer questions about the uploaded company documents.
Search documents first, then use web search to supplement. Be specific and cite sources."""),
        ("human", "{input}"),
        MessagesPlaceholder(variable_name="agent_scratchpad")
    ])
    agent = create_openai_functions_agent(llm=llm, tools=tools, prompt=prompt)
    executor = AgentExecutor(agent=agent, tools=tools, verbose=False, max_iterations=4)
    result = executor.invoke({"input": question})
    return {"answer": result["output"], "sources": ["Documents + Web Search"]}

def generate_memo() -> dict:
    if vectorstore is None:
        return {"answer": "No documents uploaded yet.", "sources": []}
    result = memo_graph.invoke({})
    return {
        "answer": result["memo"],
        "sources": ["Market Agent", "Founder Agent", "Finance Agent", "Risk Agent"]
    }
