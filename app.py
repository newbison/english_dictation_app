from flask import Flask, render_template, request, jsonify
import pandas as pd
import pyttsx4
import time
from threading import Thread, Event

app = Flask(__name__)

# Load the Excel data
df = pd.read_excel('words.xlsx')
print("=== DataFrame Info ===")
print(df.info())
print("\n=== Sample Data ===")
print(df.head())




# Global control events
stop_event = Event()
pause_event = Event()
current_word_index = 0

def speak_word(engine, word, gap):
    try:
        if stop_event.is_set():
            return False
            
        while pause_event.is_set():
            time.sleep(0.1)  # Check pause status every 100ms
            if stop_event.is_set():
                return False
            
        # First time
        print(f'Pronouncing: {word} (1st time)')
        engine.say(word)
        engine.runAndWait()
        
        if stop_event.is_set():
            return False
            
        while pause_event.is_set():
            time.sleep(0.1)
            if stop_event.is_set():
                return False
                
        time.sleep(1)  # Pause between repetitions
        
        if stop_event.is_set():
            return False
            
        while pause_event.is_set():
            time.sleep(0.1)
            if stop_event.is_set():
                return False
            
        # Second time
        print(f'Pronouncing: {word} (2nd time)')
        engine.say(word)
        engine.runAndWait()
        
        if stop_event.is_set():
            return False
            
        while pause_event.is_set():
            time.sleep(0.1)
            if stop_event.is_set():
                return False
            
        time.sleep(gap)  # Pause between words
        return True
    except Exception as e:
        print(f'Error in speak_word: {e}')
        return False

def pronounce_words(word_list):
    global current_word_index
    try:
        # Initialize engine for this session
        engine = pyttsx4.init()
        engine.setProperty('rate', 100)
        engine.setProperty('volume', 1)
        voices = engine.getProperty('voices')
        engine.setProperty('voice', voices[1].id)

        for i, word in enumerate(word_list):
            current_word_index = i
            if stop_event.is_set():
                break
                
            length = len(word.split(" "))
            if length == 1:
                gap = 2 
            elif length == 2: 
                gap = 3 
            elif length == 3: 
                gap = 5
            else: 
                gap = 9
            
            if not speak_word(engine, word, gap):
                break

        engine.stop()
        del engine
        current_word_index = 0
        return True
    except Exception as e:
        print(f'Error in pronounce_words: {e}')
        current_word_index = 0
        return False

def async_pronounce(words):
    global current_word_index
    stop_event.clear()  # Clear any previous stop signal
    pause_event.clear()  # Clear any previous pause signal
    current_word_index = 0
    pronounce_words(words)

@app.route('/')
def index():
    # Get unique values for dropdowns
    grades = sorted(df['grade'].unique().tolist())
    semesters = sorted(df['semester'].unique().tolist())
    models = sorted(df['model'].unique().tolist())
    units = sorted(df['unit'].unique().tolist())
    categories = sorted(df['category'].unique().tolist())
    
    # Set default values (first item in each list)
    default_grade = "5" if grades else ""
    default_semester = "1" if semesters else ""
    default_model = str(models[0]) if models else ""
    default_unit = str(units[0]) if units else ""
    default_category = "All Categories"  # Default to show all categories
    
    print("\n=== Unique Values ===")
    print(f"Grades: {grades}")
    print(f"Semesters: {semesters}")
    print(f"Models: {models}")
    print(f"Units: {units}")
    print(f"Categories: {categories}")
    
    return render_template('index.html', 
                         grades=grades,
                         semesters=semesters,
                         models=models,
                         units=units,
                         categories=categories,
                         default_grade=default_grade,
                         default_semester=default_semester,
                         default_model=default_model,
                         default_unit=default_unit,
                         default_category=default_category)

@app.route('/get_words', methods=['POST'])
def get_words():
    data = request.json
    grade = data.get('grade')
    semester = data.get('semester')
    mode = data.get('model')
    unit = data.get('unit')
    category = data.get('category')
    
    # Convert DataFrame columns to strings for comparison
    df['grade'] = df['grade'].astype(str)
    df['semester'] = df['semester'].astype(str)
    df['model'] = df['model'].astype(str)
    df['unit'] = df['unit'].astype(str)
    
    # Filter the dataframe based on selections
    filtered_df = df[
        (df['grade'] == str(grade)) &
        (df['semester'] == str(semester)) &
        (df['model'] == str(mode)) &
        (df['unit'] == str(unit))
    ]
    
    # Apply category filter if specific category is selected
    if category and category != "All Categories":
        filtered_df = filtered_df[filtered_df['category'] == category]
    
    print("\n=== Filtered DataFrame ===")
    print(f"Number of rows found: {len(filtered_df)}")
    print(filtered_df)
    
    # Convert filtered dataframe to dict
    words_data = filtered_df[['category', 'English']].to_dict('records')
    return jsonify(words_data)

@app.route('/play_words', methods=['POST'])
def play_words():
    data = request.json
    words = [item['English'] for item in data['words']] + ['Sophia, the dictation is over now.']
    
    # Start pronunciation in a separate thread
    thread = Thread(target=async_pronounce, args=(words,))
    thread.daemon = True
    thread.start()
    
    return jsonify({"status": "started"})

@app.route('/stop_words', methods=['POST'])
def stop_words():
    stop_event.set()  # Set the stop signal
    pause_event.clear()  # Clear pause signal when stopping
    return jsonify({"status": "stopped"})

@app.route('/pause_words', methods=['POST'])
def pause_words():
    if pause_event.is_set():
        pause_event.clear()  # Resume
        return jsonify({"status": "resumed", "current_word": current_word_index})
    else:
        pause_event.set()  # Pause
        return jsonify({"status": "paused", "current_word": current_word_index})

if __name__ == '__main__':
    app.run(debug=True)
