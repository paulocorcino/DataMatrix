function miniMenu_agenda(){ //menu incluir area

			if(document.form.titulo_agenda.value==''){
				alert("Digite o titulo do agendamento");
			}else{
				if(document.form.data_agenda.value==''){
					alert("Digite a data de agendamento");
				}else{
					document.form.method = "post"
					document.form.action = 'default.asp?op=cad_agenda&aspid=' + AspId
					document.form.submit();
				}
			}
}