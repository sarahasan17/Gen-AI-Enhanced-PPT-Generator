# Gen AI-Enhanced PPT Generator

This project generates AI-enhanced PowerPoint presentations from Word, PDF, Excel, and PPT files using the ChatGroq Llama-3.3-70B model.

## Getting Started

Follow these steps to set up and run the project.

### **Installation**
1. **Clone the Repository**
   ```bash
   git clone https://github.com/sarahasan17/Gen-AI-Enhanced-PPT-Generator.git
   cd Gen-AI-Enhanced-PPT-Generator
   ```

2. **Create a Virtual Environment (Optional but Recommended)**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On macOS/Linux
   venv\Scripts\activate    # On Windows
   ```

3. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

### **Set Up API Key**
Create a `.env` file in the project root and add your API key:
```
GROQ_API_KEY=your_api_key_here
```

Ensure that `.env` is included in `.gitignore` to prevent exposing sensitive information.

### **Run the Application**
Start the Streamlit web app:
```bash
streamlit run main.py
```

This will launch the web interface where you can upload documents and generate AI-enhanced PowerPoint presentations.

### **Contributing**
Feel free to fork the repository and submit pull requests for improvements.

### **License**
This project is licensed under the MIT License.

