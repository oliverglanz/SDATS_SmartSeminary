# Docker Deployment Guide

This folder now includes a Docker-based deployment for the Smart Seminary Streamlit app.

## What is included

- `Dockerfile`
- `docker-compose.yml`
- `.env.example`

The app image includes:

- the Streamlit application in `programs/app`
- the bundled source files in `0_source_files`

## First-time setup

From the `SmartSeminary` folder:

```bash
cp .env.example .env
```

Then edit `.env` and set strong passwords for:

- `SMART_SEMINARY_PASSWORD_KAREN`
- `SMART_SEMINARY_PASSWORD_SHARI`

For shared server use, leave this disabled:

- `SMART_SEMINARY_ENABLE_CLOSE_BUTTON=false`

## Start the app

```bash
docker compose up -d --build
```

Open:

- `http://SERVER_IP:8501`

To view logs:

```bash
docker compose logs -f
```

To stop:

```bash
docker compose down
```

## Platform compatibility

This app is packaged as a Linux container, not as separate native installers.

That means:

- macOS Apple Silicon can run it through Docker as `linux/arm64`
- Linux AMD64 servers can run it as `linux/amd64`
- Windows AMD64 machines can run it through Docker Desktop as a Linux container on `linux/amd64`

Important detail:

- `docker compose build` on a Mac ARM machine usually builds only an `arm64` image for that machine
- that local image will not automatically be reusable on AMD64 machines

If you want one publishable image name that works on all of those systems, build and push a multi-architecture image manifest.

## Multi-architecture build

Use Docker Buildx to publish both architectures from the same source tree.

Example:

```bash
docker buildx create --use --name streamline-builder
docker buildx build \
  --platform linux/amd64,linux/arm64 \
  -t YOUR_REGISTRY/streamline:latest \
  --push \
  .
```

Replace `YOUR_REGISTRY/streamline:latest` with your actual registry path, for example Docker Hub or GitHub Container Registry.

Once that image is pushed, each machine will automatically pull the correct variant for its CPU:

- Apple Silicon Mac: `linux/arm64`
- Linux AMD server: `linux/amd64`
- Windows AMD machine with Docker Desktop in Linux-container mode: `linux/amd64`

## Local builds versus shared images

Use this when you only need to run locally on the current machine:

```bash
docker compose up -d --build
```

Use a multi-arch `buildx` push when you want to distribute the same image to multiple machine types:

```bash
docker buildx build \
  --platform linux/amd64,linux/arm64 \
  -t YOUR_REGISTRY/streamline:latest \
  --push \
  .
```

Then on each target machine you can run:

```bash
docker pull YOUR_REGISTRY/streamline:latest
```

or reference that same image from `docker-compose.yml`.

## Windows note

For Windows AMD machines, the target should be Docker Desktop running Linux containers.

This setup is not intended for Windows-native containers.

## Scenario 1: Existing LAN plus optional outside access

Use the default bind address:

```env
SMART_SEMINARY_BIND_ADDRESS=0.0.0.0
```

This makes the app reachable from other machines on the same network at:

- `http://SERVER_IP:8501`

If you want outside-world access, you must also configure the network:

- allow TCP port `8501` through the Linux firewall
- forward external port `8501` on the router to the Linux server
- preferably place a reverse proxy with HTTPS in front of Streamlit

Recommended next step for internet exposure:

- put Caddy or Nginx in front of the container
- terminate HTTPS there
- optionally restrict access with VPN or IP allowlists

## Scenario 2: Dedicated Wi-Fi network

Use the same Docker setup. The difference is only how the Linux machine is networked.

Typical options:

1. Connect the Linux server to a dedicated Wi-Fi router.
2. Have users join that router's SSID.
3. Access the app using the server's IP on that dedicated network.

In this scenario, the app is still reached the same way:

- `http://SERVER_IP:8501`

Docker does not need to change. Only the surrounding network changes.

## Notes on binding

If you want the app reachable only from the server itself, use:

```env
SMART_SEMINARY_BIND_ADDRESS=127.0.0.1
```

If you want the app reachable from other devices on the network, use:

```env
SMART_SEMINARY_BIND_ADDRESS=0.0.0.0
```

## Security notes

- The current app uses simple in-app password gates, not full enterprise authentication.
- Do not expose the app to the public internet without HTTPS and stronger access controls.
- Change the default passwords before deployment.
- The Docker deployment disables the in-app `Close App` button so one user cannot shut down the shared server for everyone else.

## Useful commands

Rebuild after code changes:

```bash
docker compose up -d --build
```

Check container status:

```bash
docker compose ps
```

Open a shell in the container:

```bash
docker compose exec streamline sh
```
