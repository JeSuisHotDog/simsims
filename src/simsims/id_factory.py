_state = {
    "worker": 0,
    "food": 0,
    "product": 0
}

def get_next_worker_id():
    _state["worker"] += 1
    return _state["worker"]

def get_next_food_id():
    _state["food"] += 1
    return _state["food"]

def get_next_product_id():
    _state["product"] += 1
    return _state["product"]