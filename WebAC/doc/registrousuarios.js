// listar usuarios

function getListItemsUsuarios() {
  var url = "https://glencore.sharepoint.com/sites/co-lmn-ingenieriayproyectossib/_api/web/lists/getbytitle('Tipos_usuario')/items?$filter=Aplicacion eq 'Comprometidos'&$orderby=Created asc";

  $.ajax({
      url: url,
      type: "GET",
      headers: {
          "Accept": "application/json; odata=verbose"
      },
      success: function(data) {
          var items = data.d.results;
          var table = "<table id='myTableUsuarios' class='display'><thead><tr><th>ID</th><th>Nombre Usuario</th><th>Correo Usuario</th><th>Tipo Usuario</th><th>Acciones</th></tr></thead><tbody>";

          for (var i = 0; i < items.length; i++) {
              var id = items[i].ID;
              var Nombre = items[i].Nombre_Usuario;
              var Correo = items[i].Title;
              var Tipousuario = items[i].Tipo_usuario;

              table += "<tr><td>" + id + "</td><td>" + Nombre + "</td><td>" + Correo + "</td><td>" + Tipousuario + "</td><td><div class='btn-group'><button type='button' class='btn btn-secondary dropdown-toggle' data-bs-toggle='dropdown' aria-expanded='false'>Ver más</button><ul class='dropdown-menu bg-white border'><li><a class='dropdown-item' href='#' onclick='deleteListItemUser(" + id + ")'>Eliminar</a></li><li><a class='dropdown-item' href='#' data-bs-toggle='modal' data-bs-target='#FormUsuariosModal' onclick='getUserById(" + id + ")'>Editar</a></li></ul></div></td></tr>";
          }

          table += "</tbody></table>";
          $("#myTableContainerUsuarios").html(table);

          // Inicializar el DataTable con opciones de filtrado
          var dataTable = $('#myTableUsuarios').DataTable({
              "order": [[1, "asc"]],
              "language": {
                  "url": "//cdn.datatables.net/plug-ins/1.11.3/i18n/es_es.json"
              },
              "searching": true
          });

          // Agregar el elemento de búsqueda a cada encabezado de columna
          $('#myTableUsuarios thead th').each(function() {
              var title = $(this).text();
              $(this).html('<input type="text" placeholder="Buscar ' + title + '" />');
          });

          // Aplicar la búsqueda por columna
          dataTable.columns().every(function() {
              var that = this;

              $('input', this.header()).on('keyup change', function() {
                  if (that.search() !== this.value) {
                      that.search(this.value).draw();
                  }
              });
          });
      },
      error: function(error) {
          console.log(JSON.stringify(error));
      }
  });
}


    // Funcion Registrar USUARIOS---------------------------------------------------------------------------------------------------------
      
    function createListItemUsuarios() {
  
    
  
  
  

      var userInfo = $("#CampoUsuariosID").getUserInfo();
      var nombre = userInfo[0];
      var correo = userInfo[1];
      var correoglencoresinformato = userInfo[4];
      var correoglencoreformateado = correoglencoresinformato.replace('i:0#.f|membership|', '');

      
      var TipoUsuario2 = document.getElementById("CampoTipoUsuarioID").value;
  
     
  
  
      // Obtener el contexto del sitio del sitio padre
      var siteUrl = "https://glencore.sharepoint.com/sites/co-lmn-ingenieriayproyectossib"; // URL del sitio padre
      var ctx = new SP.ClientContext(siteUrl);
  
      // Obtener la lista en el sitio padre por su título
      var list = ctx.get_web().get_lists().getByTitle('Tipos_usuario');
  
      // Crear un nuevo elemento de lista
      var itemCreateInfo = new SP.ListItemCreationInformation();
      var listItem = list.addItem(itemCreateInfo);
  
    
      listItem.set_item('Title',  correoglencoreformateado);
      listItem.set_item('Nombre_Usuario',  nombre);
      listItem.set_item('Tipo_usuario', TipoUsuario2);
      listItem.set_item('Aplicacion', "Comprometidos");
   
  
      
      listItem.update();
      ctx.executeQueryAsync(onSuccess1, onFail1);
    }
      
    function onSuccess1() {
      getListItemsUsuarios();
      Swal.fire({
        icon: 'success',
        title: '¡Listo!',
        allowOutsideClick: false,
        text: 'Usuario registrado correctamente.',
        didClose: () => {
          var boton = document.getElementById("btncerrarmodalformUsuarios");
          boton.click();
        }
      });
    }
    
    function onFail1(sender, args) {
      alert('Error al crear la acción: ' + args.get_message());
    }
   
    //FIN  FUNCION EN MODAL DE AGREGAR USUARIOS
    
    
    

    
//---------------------------------------------------------------INICIO FUNCION ELIMINAR USUARIO-----------------------------------------------------------------------------------------------------------

  function deleteListItemUser(id) {
    
    Swal.fire({
      title: '¿Estás seguro?',
      text: 'El registro será eliminado permanentemente',
      icon: 'warning',
      showCancelButton: true,
      confirmButtonText: 'Sí, eliminar',
      cancelButtonText: 'Cancelar'
    }).then((result) => {
      if (result.isConfirmed) {

        // var clientContext = new SP.ClientContext.get_current();
        var siteUrl = "https://glencore.sharepoint.com/sites/co-lmn-ingenieriayproyectossib/";
        var clientContext = new SP.ClientContext(siteUrl);
        var oList = clientContext.get_web().get_lists().getByTitle('Tipos_usuario');
        var camlQuery = new SP.CamlQuery();
        camlQuery.set_viewXml('<View><Query><Where><Eq><FieldRef Name=\'ID\'/><Value Type=\'Text\'>' + id + '</Value></Eq></Where></Query></View>');
        var oListItems = oList.getItems(camlQuery);
        clientContext.load(oListItems);
        clientContext.executeQueryAsync(
          function() {
            var listItemEnumerator = oListItems.getEnumerator();
            while (listItemEnumerator.moveNext()) {
              var listItem = listItemEnumerator.get_current();
              listItem.deleteObject();
              clientContext.executeQueryAsync(
                function() {
                  Swal.fire({
                    icon: 'success',
                    title: '¡Listo!',
                    allowOutsideClick: false, // Permitir cerrar la alerta haciendo clic fuera de ella
                    text: 'Registro eliminado con éxito!'
                  }).then((result) => {
                    if (result.isConfirmed) {
                      getListItemsUsuarios();
                    }
                  });
                },
                function(sender, args) {
                  alert(args.get_message());
                }
              );
            }
          },
          function(sender, args) {
            alert(args.get_message());
          }
        );
      }
    });
  }
  //---------------------------------------------------------------FIN FUNCION ELIMINAR  USUARIO-----------------------------------------------------------------------------------------------------------


    //INICIO FUNCION OBTENER USUARIO POR ID ----------------------------------------------------------------------------------------------------------------------------

/**
 * Esta función recupera y muestra la información de un elemento específico de la lista de SharePoint "user_type" basándose en su ID.
 * Se utiliza  para llenar el formulario con los datos del elemento seleccionado.
 * 
 */
function getUserById(id) {
  document.getElementById("botonCrearUsuario").id = "botonActualizarUsuario";
  document.getElementById("botonActualizarUsuario").textContent = "Actualizar usuario";



  // Define la URL del sitio de SharePoint donde se encuentra la lista.
  var siteUrl = "/sites/co-lmn-ingenieriayproyectossib/";
  // Crea un contexto del cliente para interactuar con la API de SharePoint.
  var clientContext = new SP.ClientContext(siteUrl);
  // Obtiene la referencia a la lista "user_type" dentro del sitio especificado.
  var oList = clientContext.get_web().get_lists().getByTitle('Tipos_usuario');

  // Recupera el elemento por su ID.
  var listItem = oList.getItemById(id);
  // Prepara el elemento para su carga con la llamada subsiguiente.
  clientContext.load(listItem);
  // Ejecuta la consulta asíncrona para cargar el elemento.
  clientContext.executeQueryAsync(
  function () {
    // En caso de éxito, rellena el formulario HTML con los datos del elemento.

  // Establece el valor del elemento 'itemId' (El cual contiene el id del elemento en el formulario) con el ID recuperado.
  document.getElementById("itemUsersId").value = id;

  // Rellena diferentes partes del formulario con los datos del elemento.

  document.getElementById('CampoUsuariosID').innerHTML = listItem.get_item('Nombre_Usuario');
  document.getElementById('CampoTipoUsuarioID').value = listItem.get_item('Tipo_usuario');


},
function (sender, args) {
  // En caso de error, muestra un mensaje con la información del error.
  console.log(args.get_message());
}
);


 



}
// FIN FUNCION OBTENER USUARIO POR ID ----------------------------------------------------------------------------------------------------------------------------




// Función principal para actualizar el elemento en la lista
function updateListItemUser() {
 
  itemId = document.getElementById('itemUsersId').value;

  // Obtiene los valores de los campos relacionados con el proyecto y el usuario
  var TipoUsuario = document.getElementById("CampoTipoUsuarioID").value;

  // --------------------------------------------------------------------------------------------------------------------------------

  // Configuración para acceder a la lista de SharePoint
  var siteUrl = "/sites/co-lmn-ingenieriayproyectossib/";
  var clientContext = new SP.ClientContext(siteUrl);
  var oList = clientContext.get_web().get_lists().getByTitle('Tipos_usuario');

  // Obtiene el elemento de la lista por su ID
  this.oListItem = oList.getItemById(itemId);

  // Actualiza los valores de los campos del elemento
  oListItem.set_item('Tipo_usuario', TipoUsuario);

  // Actualiza el elemento en la lista de SharePoint
  oListItem.update();
  clientContext.executeQueryAsync(Function.createDelegate(this, this.onSuccess2), Function.createDelegate(this, this.onFail2));


}

// Función llamada cuando el nuevo usuario es creado exitosamente.
function onSuccess2() {
  getListItemsUsuarios();
  // Mostrar un mensaje de éxito y cerrar el modal del formulario.
  Swal.fire({
    icon: 'success',
    title: '¡Listo!',
    allowOutsideClick: false,
    text: 'Usuario actualizado correctamente.',
    didClose: () => {
      var boton = document.getElementById("btncerrarmodalformUsuarios");
      boton.click();
    }
  });
}

// Función llamada si ocurre un error al crear el nuevo usuario.
function onFail2(sender, args) {
  alert('Error al crear la acción: ' + args.get_message());
}

