Client

	- Render:
		Implementar un batch render, para optimizar el renderizado.
		Abstraer las llamadas al render de la libreria (DirectX).
		Cambiar el actual render (tipico 2d) y manejar todo por matrices (nos servirá más adelante).
		Mejorar el manejo de los shaders, y dar más opciones y facilidades para utilizarlos.
		Implementar render y animación de sprites mediante huesos.
		Implementar multitexturing en el terreno.
		Implementar luces y sombras (más bonitas que lo que hay ahora por AO).
		Normal mapping.
		
	- Graphic User Interface:
		Eliminar formularios, crear una instancia de la ventana desde el main directamente.
		Crear un sistema de GUI mediante eventos (event handlers). Añadir script tipo html (como en AOC).
		Dejar la posibilidad de templates, para añadir nuevos facilmente.
		
	- Math:
		Añadir algunas operaciones con matrices, vectores y demás.
		
	- Scene:
		Cambiar el actual mapdata por un quadtree.
		Mejorar el formato de mapas, estilo Loopzer (mediante listas de datos).
		Implementar montañas, deformación de terreno.
		Implementar meteorologia y dia / noche.
		
		
	- Collision Detection:
		Añadir Quad, Circle o Sphere (si lo vamos a pasar a 2.5d).
		Añadir algunas operaciones basicas para detección de colisiones.
		
	- Physics:
		Añadir algo de fisica, no muy complejo, abstraida para poder utilizarla tanto para particulas como para personajes.
		
Server
	
	- Scene:
		Implementar el quadtree equivalente, con chequeo de rango de visión y demas,
			se simplifica todo el tema de áreas y control de usuarios y npcs.
			
	- Core:
		Implementar la conexión sockets mediante api de windows (tengo las clases hechas).
		Implementar una clase Thread, una clase Mutex y Condition (tengo las dos primeras hechas). Armar todo el servidor con hilos.
		
		
		