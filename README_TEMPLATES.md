
# Installing Templates

This project relies on templates to function correctly.  
You need to copy the `templates.zip` archive into the `app/backend` directory of your Docker container and then extract it.

## Usage Example

1. Copy `templates.zip` into the running Docker container  
```bash
docker cp templates.zip open-webui:/app/backend/
````

2. Unzip the archive inside the container

```bash
docker exec open-webui \
    unzip /app/backend/templates.zip \
    -d /app/backend/
```
3. Verify that the templates were extracted correctly

```bash
docker exec open-webui ls /app/backend/templates
```

---

> **Note:**
>
> * Make sure the file is named `templates.zip` (not `tempaltes.zip`).
> * The Docker container must be running and named `open-webui`.
> * After extraction, templates will be available under `app/backend/templates/`.


