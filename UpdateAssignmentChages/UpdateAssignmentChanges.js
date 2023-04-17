// Function to get CSRF token
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

// Get course ID and page URL from URL
const courseIdAndPageUrlRegex = /\/courses\/(\d+)\/pages\/(.+)/;
const courseIdAndPageUrlMatch = window.location.pathname.match(courseIdAndPageUrlRegex);
const courseId = courseIdAndPageUrlMatch[1];
const pageUrl = courseIdAndPageUrlMatch[2];

// Retrieve page data from Canvas API
const getPageDataResponse = await fetch(`/api/v1/courses/${courseId}/pages/${pageUrl}`);
const pageData = await getPageDataResponse.json();

// Parse page data into HTML and extract assnt-section divs
const parser = new DOMParser();
const pageDocument = parser.parseFromString(pageData.body, 'text/html');
const assntSections = pageDocument.querySelectorAll('div.assnt-section');

// Update assignments, discussions, quizzes, etc. with content of assnt-section divs
const csrfToken = getCsrfToken();
assntSections.forEach(async (assntSection) => {
  const assntContent = assntSection.outerHTML;
  const assntLink = assntSection.nextElementSibling.querySelector('a[data-api-endpoint]');
  const assntApiEndpoint = assntLink.getAttribute('data-api-endpoint');
  const assntId = assntApiEndpoint.split('/').pop();
  const path = assntApiEndpoint.split('/').slice(-3);
  let apiEndpoint;
  switch (path[1]) {
    case 'assignments':
      apiEndpoint = `/api/v1/courses/${courseId}/assignments/${assntId}`;
      break;
    case 'discussion_topics':
      apiEndpoint = `/api/v1/courses/${courseId}/discussion_topics/${assntId}`;
      break;
    case 'quizzes':
      apiEndpoint = `/api/v1/courses/${courseId}/quizzes/${assntId}`;
      break;
    // handle other types of links here
    default:
      console.log(`Unsupported link type: ${path[1]}`);
      return;
  }
  const headers = {
    'X-CSRF-Token': csrfToken,
  };
  const requestBody = new FormData();
  requestBody.append('assignment[description]', assntContent);
  const response = await fetch(apiEndpoint, {
    method: 'PUT',
    headers: headers,
    body: requestBody,
  });
  const responseBody = await response.json();
  console.log(responseBody);
});
