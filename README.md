# Calculadora del Método Simplex

Proyecto de Investigación de Operaciones. Es una calculadora para resolver problemas de programación lineal mediante el método Simplex.

## Requisitos

- Windows, Linux o macOS.
- Python 3.12 o una versión compatible de Python 3.
- Un navegador web.

La aplicación usa bibliotecas incluidas en Python. **SciPy es opcional** y se utiliza únicamente para comparar algunos resultados con HiGHS.

## Cómo ejecutar

1. Descarga o clona el repositorio.
2. Abre una terminal en la carpeta donde se encuentra el archivo `simplex_web (1).py`.
3. Comprueba que Python esté instalado:

   ```bash
   python --version
   ```

4. Ejecuta la aplicación:

   ```bash
   python "simplex_web (1).py" --puerto 9014
   ```

5. Abre en el navegador la dirección:

   ```text
   http://127.0.0.1:9014
   ```

Si el comando `python` no funciona en Windows, prueba con `py` en su lugar.

## Cómo usarla

1. Selecciona si el problema es de maximización o minimización.
2. Ingresa los coeficientes de la función objetivo.
3. Agrega las restricciones, seleccionando su signo y escribiendo el lado derecho.
4. Pulsa **Resolver modelo** para ver las tablas, los pivotes y la solución.
5. Revisa los valores de las variables y las holguras que presenta el resultado.

La versión web admite restricciones `<=`, `>=` e igualdad, y utiliza el método de Dos Fases cuando necesita variables artificiales. La gráfica está disponible para modelos de dos variables.

## Pruebas

Si quieres ejecutar las pruebas incluidas, usa:

```bash
python -m unittest -v test_simplex_api test_simplex_fase2 test_simplex_precision
```

También puedes ejecutar las autopruebas del programa:

```bash
python "simplex_web (1).py" --probar
```

Los comandos deben ejecutarse desde la carpeta que contiene el programa y los archivos de pruebas.

## Nota

El nombre del archivo puede variar si se renombra o se descarga una versión distinta. En ese caso, ajusta el comando para que coincida con el nombre real del archivo. Si el proyecto se encuentra dentro de una subcarpeta, entra en esa carpeta antes de ejecutar los comandos.
