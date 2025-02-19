import zeep
try:
# URL del servicio SOAP
    wsdl = 'https://www.banguat.gob.gt/variables/ws/tipocambio.asmx?wsdl'

    # Crear un cliente con Zeep
    client = zeep.Client(wsdl=wsdl)

    # Hacer la solicitud para obtener el tipo de cambio
    response = client.service.TipoCambioDia()

    # Mostrar la respuesta
    # Extraer la referencia del tipo de cambio del dólar
    referencia_dolar = response['CambioDolar']['VarDolar'][0]['referencia']

    # Mostrar el mensaje
    print(f"La referencia del tipo de cambio del dólar es: {referencia_dolar}")

except Exception as e:
    print(f"Error al consumir el servicio: {e}")