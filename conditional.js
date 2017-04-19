$(document).ready(function() {
		//Inputs that determine what fields to show
		var position = $('#employee input:radio[name=billable]');
		
		//Wrappers for all fields
		var employeename = $('#employee p[name=employeename]');
		var preferredname=$('#employee p[name=preferredname]');	
		var forlocation=$('#employee p[name=location]');
		var startdate=$('#employee input:text[name=startdate]');
		var enddate=$('#employee input:text[name=enddate]');
		var billable = $('#employee select[name=billable]').parent();
		var nonbillable = $('#employee select[name=non-billable]').parent();
		var billingrate = $('#employee input:text[name=billingrate]');
		var retainerrate = $('#employee input:text[name=retainerrate]');
		var practicegroup = $('#employee select[name=practicegroup]');
		var attorneyassignment = $('#employee input:text[name=attorneyassignment]');
		var attorneyassistant = $('#employee input:text[name=attorneyassistant]');
		var staffJobtitle = $('#employee select[name=position]');
		var computertype = $('#employee fieldset[name=computer]');
		var comments = $('#employee #comments');
		var email_group  = $('#employee #emailgroups');
		var all=enddate.add(billable).add(billingrate).add(retainerrate).add(nonbillable).add(practicegroup).add(attorneyassignment).add(attorneyassistant).add(staffJobtitle).add(computertype).add(computertype).add(comments).add(email_group);
				
		billable.change(function(){
			var value=this.value;						
			all.addClass('hidden'); //hide everything and reveal as needed
			
			if (value === 'Billable'){
				billable.removeClass('hidden');
				billingrate.removeClass('hidden');
				retainerrate.removeClass('hidden');
				practicegroup.removeClass('hidden');
				attorneyassistant.removeClass('hidden');
				computertype.removeClass('hidden');
				comments.removeClass('hidden');
				email_group.removeClass('hidden');
			}
			else if (value === 'Non-Billable'){
				nonbillable.removeClass('hidden');
				attorneyassignment.removeClass('hidden');
				staffJobtitle.removeClas('hidden');
				computertype.removeClass('hidden');
				comments.removeClass('hidden');
			}		
		});	

});
