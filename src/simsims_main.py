from simsims.simsims_lib import *


def main():
    """Set up and run the SimSims Petri net simulation."""
    # global next_worker_id, next_food_id, next_product_id
    next_worker_id = 1
    next_food_id = 1
    next_product_id = 1

    network = Network()

    barracks1 = Barracks("Barracks 1")
    barracks2 = Barracks("Barracks 2")
    barracks3 = Barracks("Barracks 3")
    barracks4 = Barracks("Barracks 4")
    barracks5 = Barracks("Barracks 5")

    storage1 = Storage("Storage 1")
    storage2 = Storage("Storage 2")

    food_storage1 = FoodStorage("Food Storage 1")
    food_storage2 = FoodStorage("Food Storage 2")

    network.add_place(barracks1)
    network.add_place(barracks2)
    network.add_place(barracks3)
    network.add_place(barracks4)
    network.add_place(barracks5)
    network.add_place(storage1)
    network.add_place(storage2)
    network.add_place(food_storage1)
    network.add_place(food_storage2)

    t1 = SimpleTransition("Move B1 to B2", barracks1, barracks2, danger_level=90)
    field1 = Field("Field Work", barracks2, barracks3, food_storage1, danger_level=80)
    cafeteria1 = Cafeteria("Lunch", barracks3, food_storage1, barracks4)
    factory1 = Factory("Production", barracks4, barracks5, storage1, danger_level=80)
    home1 = Home('Rest or Birth', barracks5, storage1, barracks1)

    network.add_transition(t1)
    network.add_transition(field1)
    network.add_transition(cafeteria1)
    network.add_transition(factory1)
    network.add_transition(home1)

    print("\nAdding initial workers to Barracks 1...")
    for i in range(1, 4):
        worker = Worker(id_factory.get_next_worker_id())
        barracks1.add_item(worker)

    print("Adding initial workers to Barracks 3...")
    for i in range(1, 3):
        worker = Worker(id_factory.get_next_worker_id())
        barracks3.add_item(worker)

    print("Adding initial food to Food Storage 1...")
    for i in range(1, 10):
        food = Food(id_factory.get_next_food_id())
        food_storage1.add_item(food)

    total_count = len(network.places) + len(network.transitions)
    print(f"\nTotal places and transitions: {total_count}")
    if total_count >= 11:
        print("Requirement met (at least 11)")
    else:
        print(f"WARNING: Need at least 11, have {total_count}")

    print("\n" + "=" * 70)
    print("STARTING SIMULATION")
    print("=" * 70)

    network.run()
    network.plot_results()
    network.save_to_excel()


if __name__ == "__main__":
    main()
