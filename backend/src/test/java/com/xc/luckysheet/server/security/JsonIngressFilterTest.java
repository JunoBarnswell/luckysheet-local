package com.xc.luckysheet.server.security;

import com.fasterxml.jackson.databind.ObjectMapper;
import jakarta.servlet.FilterChain;
import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockHttpServletRequest;
import org.springframework.mock.web.MockHttpServletResponse;
import java.nio.charset.StandardCharsets;
import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

class JsonIngressFilterTest {
    private MockHttpServletRequest request(String body) {
        var request = new MockHttpServletRequest("POST", "/api/workbooks/book/queries/execute") {
            @Override public long getContentLengthLong() { return -1; } // chunked admission
        };
        request.setContentType("application/json"); request.setContent(body.getBytes(StandardCharsets.UTF_8));
        return request;
    }
    @Test void validNestedJsonReachesTheSameBinderWithItsBytesIntact() throws Exception {
        var response = new MockHttpServletResponse(); var filter = new JsonIngressFilter(new ObjectMapper());
        var body = "{\"body\":{\"rows\":[1,2,3],\"label\":\"合法请求\"}}";
        filter.doFilter(request(body), response, (req, res) -> assertEquals(body, new String(req.getInputStream().readAllBytes(), StandardCharsets.UTF_8)));
        assertEquals(200, response.getStatus());
    }
    @Test void chunkedDeepWideLongTextAndOversizedBodiesAreRejectedBeforeBinding() throws Exception {
        for (String body : java.util.List.of("[".repeat(65) + "0" + "]".repeat(65),
                "[" + "0,".repeat(JsonIngressFilter.MAX_TOKENS) + "0]",
                "{\"text\":\"" + "x".repeat(1024 * 1024 + 1) + "\"}",
                " ".repeat(JsonIngressFilter.MAX_BYTES + 1))) {
            var chain = mock(FilterChain.class); var response = new MockHttpServletResponse();
            new JsonIngressFilter(new ObjectMapper()).doFilter(request(body), response, chain);
            assertEquals(413, response.getStatus()); assertTrue(response.getContentAsString().contains("JSON_RESOURCE_LIMIT"));
            verifyNoInteractions(chain);
        }
    }
    @Test void declaredOversizeIsRejectedWithoutReadingTheBody() throws Exception {
        var request = new MockHttpServletRequest("POST", "/api/workbooks") {
            @Override public long getContentLengthLong() { return JsonIngressFilter.MAX_BYTES + 1L; }
            @Override public jakarta.servlet.ServletInputStream getInputStream() { fail("must not read"); return null; }
        };
        request.setContentType("application/problem+json"); var response = new MockHttpServletResponse(); var chain = mock(FilterChain.class);
        new JsonIngressFilter(new ObjectMapper()).doFilter(request, response, chain);
        assertEquals(413, response.getStatus()); verifyNoInteractions(chain);
    }
}
