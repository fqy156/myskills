# Cloud Upload

The uploader supports three modes and is designed to work without third-party Python packages.

## Environment Variables

- `OPS_ARCH_UPLOAD_MODE`
- `OPS_ARCH_UPLOAD_TARGET`
- `OPS_ARCH_UPLOAD_URL`
- `OPS_ARCH_UPLOAD_USERNAME`
- `OPS_ARCH_UPLOAD_PASSWORD`
- `OPS_ARCH_UPLOAD_TOKEN`

## Modes

### `copy`

Use this when the machine already has a cloud-synced directory mounted locally.

```bash
export OPS_ARCH_UPLOAD_MODE=copy
export OPS_ARCH_UPLOAD_TARGET=/mnt/cloud-drive/arch-output
```

### `webdav`

Use HTTP `PUT` with optional basic auth.

```bash
export OPS_ARCH_UPLOAD_MODE=webdav
export OPS_ARCH_UPLOAD_URL=https://example.com/remote.php/dav/files/team/arch-output/
export OPS_ARCH_UPLOAD_USERNAME=my-user
export OPS_ARCH_UPLOAD_PASSWORD=my-password
```

### `http-put`

Use HTTP `PUT` with optional bearer token.

```bash
export OPS_ARCH_UPLOAD_MODE=http-put
export OPS_ARCH_UPLOAD_URL=https://storage.example.com/upload/customer-arch.pptx
export OPS_ARCH_UPLOAD_TOKEN=my-token
```

## Explicit Invocation

```bash
python3 scripts/upload_artifact.py \
  --file /absolute/path/to/output.pptx
```

If no mode is configured, the uploader returns the local file path.
