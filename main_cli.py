#!/usr/bin/env python3
"""
main_cli.py - Interfaz de línea de comandos para generar certificados

Uso:
    python main_cli.py                    # Ejecutar flujo completo
    python main_cli.py --dry-run         # Simular sin generar archivos
    python main_cli.py --verbose         # Modo verboso
    python main_cli.py --help            # Mostrar ayuda
"""

import argparse
import logging
import sys
from pathlib import Path
from config import CertificadosConfig
from tests.test_e2e import CertificadosProcessor

def setup_logging(verbose: bool = False, log_file: str = None):
    """Configura el sistema de logging"""
    level = logging.DEBUG if verbose else logging.INFO
    
    handlers = [logging.StreamHandler(sys.stdout)]
    if log_file:
        handlers.append(logging.FileHandler(log_file))
    
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=handlers
    )

def verificar_requisitos():
    """Verifica que todos los archivos necesarios estén presentes"""
    archivos_requeridos = [
        '.env',
        'plantilla.docx'
    ]
    
    faltantes = [archivo for archivo in archivos_requeridos if not Path(archivo).exists()]
    
    if faltantes:
        print(" Archivos requeridos no encontrados:")
        for archivo in faltantes:
            print(f"   - {archivo}")
        print("\n💡 Instrucciones:")
        if '.env' in faltantes:
            print("   - Copia .env.example como .env y completa la configuración")
        if 'plantilla.docx' in faltantes:
            print("   - Coloca el archivo plantilla.docx en la raíz del proyecto")
        return False
    
    return True

def mostrar_info_configuracion(config: CertificadosConfig):
    """Muestra información de la configuración actual"""
    print("\n Configuración actual:")
    print(f"    Carpeta Certificados ID: {config.certificados_folder_id}")
    print(f"    Excel File ID: {config.excel_file_id[:20]}...")
    print(f"    Plantilla: {config.plantilla_path}")
    print(f"    Carpeta temporal: {config.temp_folder}")
    print(f"    Columna certificado: {config.columna_certificado}")
    print(f"    Columna empresa: {config.columna_empresa}")

def main():
    parser = argparse.ArgumentParser(
        description="Generador de Certificados desde Excel en OneDrive",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  python main_cli.py                    # Ejecutar flujo completo
  python main_cli.py --dry-run         # Simular sin generar archivos
  python main_cli.py --verbose         # Modo verboso con más detalles
  python main_cli.py --config          # Mostrar configuración y salir
        """
    )
    
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Simular ejecución sin generar certificados reales'
    )
    
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='Mostrar información detallada durante la ejecución'
    )
    
    parser.add_argument(
        '--config',
        action='store_true',
        help='Mostrar configuración actual y salir'
    )
    
    parser.add_argument(
        '--log-file',
        help='Archivo donde guardar los logs (por defecto: certificados_log.log)'
    )

    args = parser.parse_args()
    
    # Verificar requisitos básicos
    if not verificar_requisitos():
        return 1
    
    try:
        # Cargar configuración
        config = CertificadosConfig.from_env()
        
        # Configurar logging
        log_file = args.log_file or config.log_file
        setup_logging(args.verbose, log_file)
        
        # Mostrar configuración si se solicita
        if args.config:
            mostrar_info_configuracion(config)
            return 0
        
        # Validar configuración
        if not config.validar():
            return 1
        
        mostrar_info_configuracion(config)
        
        # Crear procesador
        processor = CertificadosProcessor()
        
        if args.dry_run:
            print("\n MODO SIMULACIÓN - No se generarán archivos reales")
            print("   (Esta funcionalidad se puede implementar más adelante)")
            return 0
        
        # Ejecutar flujo
        print("\n Iniciando generación de certificados...")
        if args.verbose:
            print(f" Revisa el archivo '{log_file}' para detalles completos")
        
        success = processor.ejecutar_flujo_completo()
        
        if success:
            print("\n ¡Proceso completado exitosamente!")
            print(" Los certificados han sido organizados por empresa en OneDrive")
            print(" El archivo Excel ha sido actualizado")
            return 0
        else:
            print("\n El proceso falló. Revisa el log para más detalles")
            print(f" Archivo de log: {log_file}")
            return 1
            
    except ValueError as e:
        print(f" Error de configuración: {e}")
        print(" Verifica tu archivo .env")
        return 1
    except KeyboardInterrupt:
        print("\n Proceso interrumpido por el usuario")
        return 130
    except Exception as e:
        print(f" Error inesperado: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())