# Control de revolventes

Actualiza el `CONTROL REVOLVENTES` cada vez que llega un nuevo "Informe de Vencimientos de Credito" (dias 1 y 26).

```bash
pip install openpyxl pandas xlrd
python3 actualizar_control.py "CONTROL REVOLVENTES - anterior.xlsx" "recordatorios_pago_nuevo.xls" -o "CONTROL REVOLVENTES - <periodo>.xlsx"
```

- Rehace RESUMEN, VENCEN - ACCION, Vigentes (cap=0) y las 4 hojas de ejecutivo.
- No toca CONTACTOS EMPRESAS ni las hojas de NOTAS / EXTENSIONES (solo agrega empresas nuevas al final de CONTACTOS).
- Conserva lo capturado a mano (Intereses pagados, Forma de renovacion, Constancia, Estatus, Notas) si el credito vuelve a aparecer con la misma fecha.
- En cada hoja de ejecutivo: arriba los revolventes que vencen (capital a pagar != 0); abajo, en OTROS CREDITOS, los revolventes con capital 0 y los demas productos.
- Columna "Contacto" tomada de CONTACTOS EMPRESAS.
