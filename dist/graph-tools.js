import logger from './logger.js';
import { api } from './generated/client.js';
import { z } from 'zod';
export function registerGraphTools(server, graphClient, readOnly = false, enabledToolsPattern) {
    let enabledToolsRegex;
    if (enabledToolsPattern) {
        try {
            enabledToolsRegex = new RegExp(enabledToolsPattern, 'i');
            logger.info(`Tool filtering enabled with pattern: ${enabledToolsPattern}`);
        }
        catch (error) {
            logger.error(`Invalid tool filter regex pattern: ${enabledToolsPattern}. Ignoring filter.`);
        }
    }
    for (const tool of api.endpoints) {
        if (readOnly && tool.method.toUpperCase() !== 'GET') {
            logger.info(`Skipping write operation ${tool.alias} in read-only mode`);
            continue;
        }
        if (enabledToolsRegex && !enabledToolsRegex.test(tool.alias)) {
            logger.info(`Skipping tool ${tool.alias} - doesn't match filter pattern`);
            continue;
        }
        const paramSchema = {};
        if (tool.parameters && tool.parameters.length > 0) {
            for (const param of tool.parameters) {
                if (param.type === 'Body' && param.schema) {
                    paramSchema[param.name] = z.union([z.string(), param.schema]);
                }
                else {
                    paramSchema[param.name] = param.schema || z.any();
                }
            }
        }
        if (tool.method.toUpperCase() === 'GET' && tool.path.includes('/')) {
            paramSchema['fetchAllPages'] = z
                .boolean()
                .describe('Automatically fetch all pages of results')
                .optional();
        }
        server.tool(tool.alias, tool.description ?? '', paramSchema, {
            title: tool.alias,
            readOnlyHint: tool.method.toUpperCase() === 'GET',
        }, async (params, extra) => {
            logger.info(`Tool ${tool.alias} called with params: ${JSON.stringify(params)}`);
            try {
                logger.info(`params: ${JSON.stringify(params)}`);
                const parameterDefinitions = tool.parameters || [];
                let path = tool.path;
                const queryParams = {};
                const headers = {};
                let body = null;
                for (let [paramName, paramValue] of Object.entries(params)) {
                    // Skip pagination control parameter - it's not part of the Microsoft Graph API - I think 🤷
                    if (paramName === 'fetchAllPages') {
                        continue;
                    }
                    // Ok, so, MCP clients (such as claude code) doesn't support $ in parameter names,
                    // and others might not support __, so we strip them in hack.ts and restore them here
                    const odataParams = [
                        'filter',
                        'select',
                        'expand',
                        'orderby',
                        'skip',
                        'top',
                        'count',
                        'search',
                        'format',
                    ];
                    const fixedParamName = odataParams.includes(paramName.toLowerCase())
                        ? `$${paramName.toLowerCase()}`
                        : paramName;
                    const paramDef = parameterDefinitions.find((p) => p.name === paramName);
                    if (paramDef) {
                        switch (paramDef.type) {
                            case 'Path':
                                path = path
                                    .replace(`{${paramName}}`, encodeURIComponent(paramValue))
                                    .replace(`:${paramName}`, encodeURIComponent(paramValue));
                                break;
                            case 'Query':
                                queryParams[fixedParamName] = `${paramValue}`;
                                break;
                            case 'Body':
                                if (typeof paramValue === 'string') {
                                    try {
                                        body = JSON.parse(paramValue);
                                    }
                                    catch (e) {
                                        body = paramValue;
                                    }
                                }
                                else {
                                    body = paramValue;
                                }
                                break;
                            case 'Header':
                                headers[fixedParamName] = `${paramValue}`;
                                break;
                        }
                    }
                    else if (paramName === 'body') {
                        if (typeof paramValue === 'string') {
                            try {
                                body = JSON.parse(paramValue);
                            }
                            catch (e) {
                                body = paramValue;
                            }
                        }
                        else {
                            body = paramValue;
                        }
                        logger.info(`Set legacy body param: ${JSON.stringify(body)}`);
                    }
                }
                if (Object.keys(queryParams).length > 0) {
                    const queryString = Object.entries(queryParams)
                        .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
                        .join('&');
                    path = `${path}${path.includes('?') ? '&' : '?'}${queryString}`;
                }
                const options = {
                    method: tool.method.toUpperCase(),
                    headers,
                };
                if (options.method !== 'GET' && body) {
                    options.body = typeof body === 'string' ? body : JSON.stringify(body);
                }
                const isProbablyMediaContent = tool.errors?.some((error) => error.description === 'Retrieved media content') ||
                    path.endsWith('/content');
                if (isProbablyMediaContent) {
                    options.rawResponse = true;
                }
                logger.info(`Making graph request to ${path} with options: ${JSON.stringify(options)}`);
                let response = await graphClient.graphRequest(path, options);
                const fetchAllPages = params.fetchAllPages === true;
                if (fetchAllPages && response && response.content && response.content.length > 0) {
                    try {
                        let combinedResponse = JSON.parse(response.content[0].text);
                        let allItems = combinedResponse.value || [];
                        let nextLink = combinedResponse['@odata.nextLink'];
                        let pageCount = 1;
                        while (nextLink) {
                            logger.info(`Fetching page ${pageCount + 1} from: ${nextLink}`);
                            const url = new URL(nextLink);
                            const nextPath = url.pathname.replace('/v1.0', '');
                            const nextOptions = { ...options };
                            const nextQueryParams = {};
                            for (const [key, value] of url.searchParams.entries()) {
                                nextQueryParams[key] = value;
                            }
                            nextOptions.queryParams = nextQueryParams;
                            const nextResponse = await graphClient.graphRequest(nextPath, nextOptions);
                            if (nextResponse && nextResponse.content && nextResponse.content.length > 0) {
                                const nextJsonResponse = JSON.parse(nextResponse.content[0].text);
                                if (nextJsonResponse.value && Array.isArray(nextJsonResponse.value)) {
                                    allItems = allItems.concat(nextJsonResponse.value);
                                }
                                nextLink = nextJsonResponse['@odata.nextLink'];
                                pageCount++;
                                if (pageCount > 100) {
                                    logger.warn(`Reached maximum page limit (100) for pagination`);
                                    break;
                                }
                            }
                            else {
                                break;
                            }
                        }
                        combinedResponse.value = allItems;
                        if (combinedResponse['@odata.count']) {
                            combinedResponse['@odata.count'] = allItems.length;
                        }
                        delete combinedResponse['@odata.nextLink'];
                        response.content[0].text = JSON.stringify(combinedResponse);
                        logger.info(`Pagination complete: collected ${allItems.length} items across ${pageCount} pages`);
                    }
                    catch (e) {
                        logger.error(`Error during pagination: ${e}`);
                    }
                }
                if (response && response.content && response.content.length > 0) {
                    const responseText = response.content[0].text;
                    const responseSize = responseText.length;
                    logger.info(`Response size: ${responseSize} characters`);
                    try {
                        const jsonResponse = JSON.parse(responseText);
                        if (jsonResponse.value && Array.isArray(jsonResponse.value)) {
                            logger.info(`Response contains ${jsonResponse.value.length} items`);
                            if (jsonResponse.value.length > 0 && jsonResponse.value[0].body) {
                                logger.info(`First item has body field with size: ${JSON.stringify(jsonResponse.value[0].body).length} characters`);
                            }
                        }
                        if (jsonResponse['@odata.nextLink']) {
                            logger.info(`Response has pagination nextLink: ${jsonResponse['@odata.nextLink']}`);
                        }
                        const preview = responseText.substring(0, 500);
                        logger.info(`Response preview: ${preview}${responseText.length > 500 ? '...' : ''}`);
                    }
                    catch (e) {
                        const preview = responseText.substring(0, 500);
                        logger.info(`Response preview (non-JSON): ${preview}${responseText.length > 500 ? '...' : ''}`);
                    }
                }
                // Convert McpResponse to CallToolResult with the correct structure
                const content = response.content.map((item) => {
                    // GraphClient only returns text content items, so create proper TextContent items
                    const textContent = {
                        type: 'text',
                        text: item.text,
                    };
                    return textContent;
                });
                const result = {
                    content,
                    _meta: response._meta,
                    isError: response.isError,
                };
                return result;
            }
            catch (error) {
                logger.error(`Error in tool ${tool.alias}: ${error.message}`);
                const errorContent = {
                    type: 'text',
                    text: JSON.stringify({
                        error: `Error in tool ${tool.alias}: ${error.message}`,
                    }),
                };
                return {
                    content: [errorContent],
                    isError: true,
                };
            }
        });
    }
}
