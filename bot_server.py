from fastapi import FastAPI, Query
from sentence_transformers import SentenceTransformer
import faiss
import numpy as np
import uvicorn

app = FastAPI(title="Search Engine API")

# --- Loading resources (ONLY ONCE at startup) ---
print("Ładowanie modelu i indeksu... Proszę czekać.")
MODEL_PATH = r'C:\Model\all-MiniLM-L6-v2'
model = SentenceTransformer(MODEL_PATH)
index = faiss.read_index("faiss_index.bin") # Loading a file into RAM
ids_mapping = np.load("faiss_mapping_ids.npy", allow_pickle=True) # ID Mapping (your extracted case numbers)
print("Serwer gotowy!")

@app.get("/search")
def search(q: str = Query(None, title="Query string")):
    if not q:
        return {"results": []}

    query_vector = model.encode([q]).astype('float32') # Query vectorization

    distances, indices = index.search(query_vector, k=5) # FAISS search (k=5 best results)

# --- Extracting case's IDs from the mapping ---
    found_ids = []
    for idx in indices[0]:
        if idx != -1: # FAISS returns -1 if nothing is found
            found_ids.append(str(ids_mapping[idx]))

    return {
        "query": q,
        "results": found_ids  # list for Power Apps
    }

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000) # Running the server on the IP "under the mask" and port 8000