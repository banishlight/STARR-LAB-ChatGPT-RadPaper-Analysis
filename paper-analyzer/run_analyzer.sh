#!/bin/bash
# run_analyzer.sh — Start Ollama, run the paper analyzer, clean up.

# Usage: ./run_analyzer.sh [model] [extra analyzer args, e.g. --figures]
MODEL="${1:-qwen2.5:14b}"
[ $# -gt 0 ] && shift

echo "Starting Ollama..."
ollama serve 2>/dev/null &
OLLAMA_PID=$!
sleep 2

# Make sure it's actually running
if ! kill -0 $OLLAMA_PID 2>/dev/null; then
    echo "Error: Ollama failed to start"
    exit 1
fi

echo "Running analyzer with model: $MODEL"
python3 rad_paper_analyzer_local.py --model "$MODEL" "$@"

echo "Stopping Ollama..."
kill $OLLAMA_PID 2>/dev/null
wait $OLLAMA_PID 2>/dev/null
echo "Done. VRAM freed."