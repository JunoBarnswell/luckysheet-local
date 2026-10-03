package com.xc.luckysheet.server.security;

import com.fasterxml.jackson.core.JsonFactory;
import com.fasterxml.jackson.core.StreamReadConstraints;
import com.fasterxml.jackson.core.exc.StreamConstraintsException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.contract.ApiErrorResponse;
import jakarta.servlet.*;
import jakarta.servlet.http.*;
import org.springframework.core.Ordered;
import org.springframework.core.annotation.Order;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Component;
import org.springframework.web.filter.OncePerRequestFilter;
import java.io.*;

/** Validates bytes and streaming tokens before Spring can allocate nested DTOs.
 * All API JSON bodies (including unknown-length/chunked requests) share this
 * admission boundary. Semantic workbook validation still runs after binding. */
@Component
@Order(Ordered.HIGHEST_PRECEDENCE + 10)
public final class JsonIngressFilter extends OncePerRequestFilter {
    public static final int MAX_BYTES = 16 * 1024 * 1024;
    public static final int MAX_TOKENS = 2_000_000;
    private final ObjectMapper mapper;
    private static final JsonFactory FACTORY = JsonFactory.builder().streamReadConstraints(StreamReadConstraints.builder()
            .maxNestingDepth(64).maxStringLength(1024 * 1024).maxNameLength(1024)
            .maxNumberLength(256).maxDocumentLength(MAX_BYTES).build()).build();
    public JsonIngressFilter(ObjectMapper mapper) { this.mapper = mapper; }

    @Override protected void doFilterInternal(HttpServletRequest request, HttpServletResponse response, FilterChain chain)
            throws IOException, ServletException {
        String contentType = request.getContentType();
        boolean json = false;
        try { if (contentType != null) json = MediaType.APPLICATION_JSON.isCompatibleWith(MediaType.parseMediaType(contentType))
                || MediaType.parseMediaType(contentType).getSubtype().endsWith("+json"); }
        catch (IllegalArgumentException invalid) { /* Spring reports unsupported content types. */ }
        if (!request.getRequestURI().startsWith(request.getContextPath() + "/api/") || !json) { chain.doFilter(request, response); return; }
        byte[] body;
        try {
            if (request.getContentLengthLong() > MAX_BYTES) throw new BudgetExceeded();
            body = request.getInputStream().readNBytes(MAX_BYTES + 1);
            if (body.length > MAX_BYTES) throw new BudgetExceeded();
            validate(body);
        } catch (BudgetExceeded | StreamConstraintsException limit) {
            response.setStatus(413); response.setContentType(MediaType.APPLICATION_JSON_VALUE);
            mapper.writeValue(response.getOutputStream(), new ApiErrorResponse("JSON_RESOURCE_LIMIT", "JSON request exceeds its byte, depth, text or token budget; reduce the request before retrying"));
            return;
        } catch (com.fasterxml.jackson.core.JsonProcessingException malformed) {
            response.setStatus(400); response.setContentType(MediaType.APPLICATION_JSON_VALUE);
            mapper.writeValue(response.getOutputStream(), new ApiErrorResponse("VALIDATION_ERROR", "Request JSON is invalid"));
            return;
        }
        chain.doFilter(new HttpServletRequestWrapper(request) {
            @Override public ServletInputStream getInputStream() {
                ByteArrayInputStream input = new ByteArrayInputStream(body);
                return new ServletInputStream() {
                    @Override public int read() { return input.read(); }
                    @Override public int read(byte[] b, int off, int len) { return input.read(b, off, len); }
                    @Override public boolean isFinished() { return input.available() == 0; }
                    @Override public boolean isReady() { return true; }
                    @Override public void setReadListener(ReadListener listener) { throw new IllegalStateException("Synchronous JSON binding required"); }
                };
            }
            @Override public BufferedReader getReader() { return new BufferedReader(new InputStreamReader(getInputStream(), java.nio.charset.StandardCharsets.UTF_8)); }
        }, response);
    }
    public static void validate(byte[] body) throws IOException {
        if (body.length > MAX_BYTES) throw new BudgetExceeded();
        try (var parser = FACTORY.createParser(body)) {
            int tokens = 0;
            long text = 0;
            while (parser.nextToken() != null) {
                if (++tokens > MAX_TOKENS) throw new BudgetExceeded();
                if (parser.currentToken() == com.fasterxml.jackson.core.JsonToken.VALUE_STRING
                        || parser.currentToken() == com.fasterxml.jackson.core.JsonToken.FIELD_NAME) {
                    int length = parser.getTextLength();
                    if (length > (parser.currentToken() == com.fasterxml.jackson.core.JsonToken.FIELD_NAME ? 1024 : 1024 * 1024)) throw new BudgetExceeded();
                    text += length;
                }
                if (text > MAX_BYTES) throw new BudgetExceeded();
            }
        }
    }
    public static final class BudgetExceeded extends IOException {}
}
