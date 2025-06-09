import requests
import random
import time

def submit_random_entries(form_id=2756, runs=100):
    base_url = "https://datacapture-ws-public-europe.monterosa.cloud"
    submit_url = f"{base_url}/submit/{form_id}"


    # Generate a list of 100 placeholder names
    names = ["James", "John", "Robert", "Michael", "William", "David", "Richard", "Joseph", 
               "Charles", "Thomas", "Mary", "Patricia", "Jennifer", "Linda", "Elizabeth", 
               "Barbara", "Susan", "Jessica", "Sarah", "Karen", "Emily", "Emma", "Olivia", 
               "Ava", "Isabella", "Sophia", "Mia", "Charlotte", "Amelia", "Harper"]

    # A few example explanations to choose from
    explanations = [
        "He is a superstar defender.",
        "His interceptions are top-notch.",
        "Solid presence at the back.",
        "Leadership skills on the pitch.",
        "Consistent performance game to game.",
        "Exceptional tackling ability.",
        "Great reading of the game.",
        "Always positions himself well.",
        "Impeccable timing on challenges.",
        "Defensive rock for the team."
    ]

    for i in range(1, runs + 1):
        name = random.choice(names)
        explanation = random.choice(explanations)
        payload = {
            "29100": name,
            "29101": "Trent Arnold",
            "29102": explanation
        }

        response = requests.post(submit_url, json=payload)
        print(f"Run {i:03d}: name={name}, status={response.status_code}")

        # Pause briefly to avoid overwhelming the server
        time.sleep(random.uniform(0.5, 1.5))

if __name__ == "__main__":
    submit_random_entries(runs=1000)
