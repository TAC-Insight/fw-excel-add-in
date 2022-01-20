/* eslint-disable no-undef */

export function createOrUpdateQuery(queryName, url, apiKey, query) {
  localStorage.setItem(
    queryName,
    JSON.stringify({
      url,
      apiKey,
      query,
    })
  );
}

export function getQuery(queryName) {
  return JSON.parse(localStorage.getItem(queryName));
}

export function listQueries() {
  let queries = Object.keys(localStorage);
  queries = queries.filter((queryName) => {
    return queryName !== "Office API client";
  });
  return queries;
}

export function deleteQuery(queryName) {
  localStorage.removeItem(queryName);
}
