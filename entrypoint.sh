#!/bin/bash

# --- CORREÇÃO DE DNS ---
# Garante que o arquivo de DNS do contêiner seja sobrescrito corretamente.
# Usa Cloudflare (1.1.1.1) e Google (8.8.8.8).
echo -e "nameserver 1.1.1.1\nnameserver 8.8.8.8" > /etc/resolv.conf

# Aguarda 2 segundos para evitar problemas de timing no DNS
sleep 2

# Executa o comando principal do container
exec "$@"
