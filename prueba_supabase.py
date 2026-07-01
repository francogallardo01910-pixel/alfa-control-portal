from supabase import create_client

# NUEVO PROYECTO
SUPABASE_URL = "https://dpejdcfmovrjvshksdlq.supabase.co"

SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRwZWpkY2Ztb3ZyanZzaGtzZGxxIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODI3NjUxMTQsImV4cCI6MjA5ODM0MTExNH0.LQJEGDUpyZZ4Yw7P0s0XFhlL6OzHOXfygJuOSSG1MtM"

print("Conectando...")

try:

    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

    respuesta = (
        supabase
        .table("trabajadores")
        .select("*")
        .execute()
    )

    print("===================================")
    print("✅ CONEXIÓN EXITOSA")
    print(respuesta.data)

except Exception as e:

    print("===================================")
    print("❌ ERROR")
    print(type(e))
    print(e)