@echo off
:: Wrapper to bypass ExecutionPolicy and pass file arguments
PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0PPT2PDF.ps1" %*