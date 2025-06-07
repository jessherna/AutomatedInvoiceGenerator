## Docker

Build the image and run your test suite in a container:

```bash
docker build -t invoice-generator .
docker run --rm invoice-generator

Or, if using Docker Compose:

bash
Copy
Edit
docker-compose up --build app
```