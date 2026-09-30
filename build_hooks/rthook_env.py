# Hook eseguito all'avvio dell'eseguibile: disattiva eventuali plugin di pydantic
# installati nell'ambiente di build (non servono all'app e possono impedirne l'avvio).
import os
os.environ.setdefault("PYDANTIC_DISABLE_PLUGINS", "__all__")
