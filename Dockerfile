# # Use the official lightweight Python image
# FROM python:3-slim

# # Set the working directory in the container
# WORKDIR /app

# # Copy and install Python dependencies
# COPY requirements.txt .
# RUN pip install --no-cache-dir -r requirements.txt

# # Copy the rest of the application code to the container
# COPY . .

# # Expose the port that Cloud Run will run on (Cloud Run will respect this but uses $PORT)
# EXPOSE 8080

# # Run the Streamlit app
# CMD streamlit run --server.port $PORT --server.enableCORS=false app.py
# Use an official Python runtime as a parent image
FROM python:3.9-slim

# Set the working directory in the container
WORKDIR /app

# Copy the current directory contents into the container at /app
COPY . /app

# Install any needed packages specified in requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Make port 8501 available to the world outside this container
EXPOSE 8501

# Run streamlit when the container launches
CMD ["streamlit", "run", "app.py"]
