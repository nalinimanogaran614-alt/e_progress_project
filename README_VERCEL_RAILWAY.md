# E-Progress Card Setup

The Flask application is independent of XAMPP. MySQL can run locally in Docker, and the same environment-variable configuration works in deployment.

## Local setup

1. Install Python 3.10+ and Docker Desktop.
2. Create the local environment file:

	```powershell
	Copy-Item .env.example .env
	```

3. Install Python packages:

	```powershell
	py -m pip install -r requirements.txt
	```

4. Start MySQL and load the schema:

	```powershell
	docker compose up -d db
	```

5. Start Flask:

	```powershell
	py app.py
	```

Open `http://127.0.0.1:5000/`.

The database schema is imported automatically the first time the Docker volume is created. To recreate an empty database after changing the schema, run `docker compose down -v` and then `docker compose up -d db`.

## Git

Commit `.env.example`, but never commit `.env`. The repository ignores local secrets, virtual environments, Python caches, logs, and Vercel build output.

## Vercel + Railway

Use Vercel with Framework `Other`, blank build command, blank output directory, and root directory `.`. Add these production environment variables in Vercel:

```text
DB_HOST=Railway public TCP host
DB_PORT=Railway public TCP port
DB_NAME=Railway MYSQLDATABASE
DB_USER=Railway MYSQLUSER
DB_PASSWORD=Railway MYSQLPASSWORD
SECRET_KEY=a-long-random-secret
```

Alternatively, set `MYSQL_PUBLIC_URL` to the Railway public connection URL. Do not use `mysql.railway.internal` from Vercel. Redeploy after changing environment variables.
