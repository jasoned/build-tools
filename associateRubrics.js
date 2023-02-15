var currentURL = window.location.href;
var courseID = /\d+/.exec(window.location.pathname)[0];
console.log("The course ID is: " + courseID);

var assignmentIds = [];
$('.assignment').each(function() {
  assignmentIds.push($(this).find('.ig-row').attr('data-item-id'));
});
console.log(assignmentIds);

var csrfToken = getCsrfToken();

$.ajax({
  type: "GET",
  url: "/api/v1/courses/" + courseID + "/rubrics",
  success: function(rubrics) {
    console.log(rubrics);
    $('.ig-row').each(function(index) {
      var select = $('<select style="margin-right: 10px;"></select>');
      select.append('<option value="">Select Rubric</option>');
      for (var i = 0; i < rubrics.length; i++) {
        select.append('<option value="' + rubrics[i].id + '">' + rubrics[i].title + ' (' + rubrics[i].points_possible + ' points possible) </option>');
      }
      $(this).append(select);
      var btn = $('<button class="associateBtn" style="margin-left: 10px;">Associate</button>');
      btn.click(function() {
        var selectedRubricId = select.val();
        if (!selectedRubricId) {
          console.error('Please select a rubric to link');
          return;
        }
        $.ajax({
          type: 'POST',
          url: '/api/v1/courses/' + courseID + '/rubric_associations',
          contentType: 'application/json',
          headers: {
            'X-XSRF-TOKEN': csrfToken
          },
          data: JSON.stringify({
            'rubric_association': {
              'rubric_id': selectedRubricId,
              'association_id': assignmentIds[index],
              'association_type': 'Assignment',
              'use_for_grading': 'true',
              'purpose': 'grading'
            }
          }),
          success: function(data) {
            console.log(data);
            btn.text("Rubric Linked");
            btn.attr("disabled", true);
            btn.css("background-color", "#ddd");
          },
          error: function(error) {
            console.error(error);
          }
        });
      });
      $(this).append(btn);
    });
  },
  error: function(error) {
    console.error(error);
  }
});


function getCsrfToken() {
  const csrfRegex = new RegExp('^_csrf_token=(.*)$');
  let csrf;
  const cookies = document.cookie.split(';');
  for (let i = 0; i < cookies.length; i++) {
    const cookie = cookies[i].trim();
    const match = csrfRegex.exec(cookie);
    if (match) {
      csrf = decodeURIComponent(match[1]);
	break;
	}
}
return csrf;
}
