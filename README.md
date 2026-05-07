# kms-website

Source for [kms.viktorbarzin.me](https://kms.viktorbarzin.me) — a single-page reference for activating
Microsoft Volume License products against the home-lab KMS host (`kms.viktorbarzin.lan` / `kms.viktorbarzin.me`).

## Stack

- **Hugo** static site (one custom layout, single `_index`)
- **YAML data file** at `data/products.yaml` is the source of truth for all GVLK tables
- **nginx:alpine** Docker image (multi-stage build via Hugo)
- **Woodpecker CI** builds + pushes to `forgejo.viktorbarzin.me/viktor/kms-website`
  and rolls the `kms-web-page` Deployment in the `kms` namespace
- **Terraform** in `infra/stacks/kms` consumes the image (`var.image_tag`)

## Local dev

```sh
hugo server -D
# → http://localhost:1313
```

## Update GVLKs

Edit `data/products.yaml`. Push. CI rebuilds and rolls.

Sources of truth for keys:

- Windows: <https://learn.microsoft.com/en-us/windows-server/get-started/kms-client-activation-keys>
- Office / Project / Visio: <https://learn.microsoft.com/en-us/deployoffice/vlactivation/gvlks>
