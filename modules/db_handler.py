import sqlite3

def obtener_bancos(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT nombre_banco FROM Banco")
    bancos = [row[0] for row in cursor.fetchall()]
    conn.close()
    return bancos

def obtener_referencias(db_path, nombre_banco):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Obtener el ID del banco
    cursor.execute("SELECT id_banco FROM Banco WHERE nombre_banco = ?", (nombre_banco,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        raise ValueError(f"No se encontró el banco: {nombre_banco}")
    id_banco = row[0]

    # Obtener las referencias y sus colores asociadas a ese banco
    cursor.execute("""
        SELECT Referencia.codigo, Referencia.observacion, Color.hex
        FROM Referencia
        JOIN Color ON Referencia.id_color = Color.id_color
        WHERE Color.id_banco = ?
    """, (id_banco,))
    resultados = cursor.fetchall()
    conn.close()

    # Devolver como dos diccionarios separados
    referencias = {
        'codigo': {},
        'observacion': {}
    }

    for codigo, observacion, hex_color in resultados:
        if codigo:
            referencias['codigo'][codigo] = hex_color
        if observacion:
            referencias['observacion'][observacion] = hex_color

    return referencias
