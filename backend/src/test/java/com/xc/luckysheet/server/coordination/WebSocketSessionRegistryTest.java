package com.xc.luckysheet.server.coordination;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.contract.CommittedOperationEnvelope;
import com.xc.luckysheet.server.contract.CommittedOperationMutation;
import com.xc.luckysheet.server.contract.OperationEnvelope;
import com.xc.luckysheet.server.contract.OperationOrigin;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookRole;
import com.xc.luckysheet.server.service.AccessControlService;
import com.xc.luckysheet.server.service.AccessProjectionService;
import com.xc.luckysheet.server.service.RangeAccessService;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;
import org.springframework.web.socket.CloseStatus;
import org.springframework.web.socket.WebSocketSession;

import java.security.Principal;
import java.time.Instant;
import java.util.List;
import java.util.Map;

import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.doThrow;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;
import static org.mockito.Mockito.never;
import static org.mockito.Mockito.times;
import static org.mockito.ArgumentMatchers.any;
import static org.junit.jupiter.api.Assertions.assertEquals;
import org.mockito.ArgumentCaptor;
import org.springframework.web.socket.TextMessage;

class WebSocketSessionRegistryTest {
    private Principal authenticated(String subject, Instant expiresAt) {
        var token = org.springframework.security.oauth2.jwt.Jwt.withTokenValue("test-token")
                .header("alg", "none").subject(subject).issuedAt(expiresAt.minusSeconds(120)).expiresAt(expiresAt).build();
        return new org.springframework.security.oauth2.server.resource.authentication.JwtAuthenticationToken(token);
    }

    @Test
    void expiredCredentialCannotReceiveCalculationNotificationsEvenWhenAclAllows() throws Exception {
        var access = mock(AccessControlService.class);
        var registry = new WebSocketSessionRegistry(new ObjectMapper(), access,
                mock(RangeAccessService.class), mock(AccessProjectionService.class));
        var session = mock(WebSocketSession.class);
        when(session.getPrincipal()).thenReturn(authenticated("reader", Instant.now().minusSeconds(1)));
        when(session.isOpen()).thenReturn(true);
        registry.subscribeCalculation("root", java.util.Set.of("leaf"), session);
        registry.broadcastAccessChanged("leaf", 7);
        verify(session, never()).sendMessage(any());
        verify(session).close(eq(CloseStatus.POLICY_VIOLATION));
        org.mockito.Mockito.verifyNoInteractions(access);
    }

    @Test
    void calculationNotificationsCarryOnlyIdentifiersAndRecheckRootPermission() throws Exception {
        AccessControlService access = mock(AccessControlService.class);
        ObjectMapper mapper = new ObjectMapper().findAndRegisterModules();
        WebSocketSessionRegistry registry = new WebSocketSessionRegistry(mapper, access,
                mock(RangeAccessService.class), mock(AccessProjectionService.class));
        WebSocketSession session = mock(WebSocketSession.class);
        when(session.getPrincipal()).thenReturn(authenticated("reader", Instant.now().plusSeconds(60)));
        when(session.isOpen()).thenReturn(true);
        when(session.getAttributes()).thenReturn(new java.util.concurrent.ConcurrentHashMap<>());
        registry.subscribeCalculation("root", java.util.Set.of("leaf", "middle"), session);
        registry.broadcastAccessChanged("leaf", 7);
        var message = ArgumentCaptor.forClass(TextMessage.class);
        verify(session).sendMessage(message.capture());
        assertEquals(mapper.readTree("{\"type\":\"calculation.changed\",\"unitId\":\"root\",\"sourceUnitId\":\"leaf\"}"), mapper.readTree(message.getValue().getPayload()));
        verify(access).require("root", "reader", WorkbookRole.VIEWER);
        doThrow(ServiceException.forbidden("Root revoked")).when(access).require("root", "reader", WorkbookRole.VIEWER);
        registry.broadcastAccessChanged("middle", 8);
        verify(session).close(eq(CloseStatus.POLICY_VIOLATION));
        registry.broadcastAccessChanged("leaf", 9);
        verify(session, times(1)).sendMessage(any());
    }

    @Test
    void calculationSubscriptionRetirementStopsSourceNotifications() throws Exception {
        WebSocketSessionRegistry registry = new WebSocketSessionRegistry(new ObjectMapper(), mock(AccessControlService.class),
                mock(RangeAccessService.class), mock(AccessProjectionService.class));
        WebSocketSession session = mock(WebSocketSession.class);
        when(session.isOpen()).thenReturn(true);
        registry.subscribeCalculation("root", java.util.Set.of("leaf"), session);
        registry.unsubscribeCalculation(session);
        registry.broadcastAccessChanged("leaf", 10);
        verify(session, never()).sendMessage(any());
    }

    @Test
    void lifecycleNotificationKeepsDependentSubscriptionAndClosesTheSourceSession() throws Exception {
        AccessControlService access = mock(AccessControlService.class);
        WebSocketSessionRegistry registry = new WebSocketSessionRegistry(new ObjectMapper(), access,
                mock(RangeAccessService.class), mock(AccessProjectionService.class));
        WebSocketSession dependent = mock(WebSocketSession.class), source = mock(WebSocketSession.class);
        for (WebSocketSession session : List.of(dependent, source)) {
            when(session.isOpen()).thenReturn(true);
            when(session.getPrincipal()).thenReturn(authenticated("reader", Instant.now().plusSeconds(60)));
            when(session.getAttributes()).thenReturn(new java.util.concurrent.ConcurrentHashMap<>());
        }
        registry.subscribeCalculation("root", java.util.Set.of("leaf"), dependent); registry.join("leaf", source);
        registry.broadcastLifecycleChanged("leaf", "trashed"); registry.broadcastLifecycleChanged("leaf", "purged");
        verify(dependent, times(2)).sendMessage(any()); verify(dependent, never()).close(any());
        verify(source, times(2)).close(eq(CloseStatus.POLICY_VIOLATION));
        var order = org.mockito.Mockito.inOrder(source);
        order.verify(source).sendMessage(any()); order.verify(source).close(eq(CloseStatus.POLICY_VIOLATION));
        var messages = ArgumentCaptor.forClass(TextMessage.class); verify(source, times(2)).sendMessage(messages.capture());
        ObjectMapper mapper = new ObjectMapper();
        assertEquals(mapper.readTree("{\"type\":\"workbook.lifecycle.changed\",\"unitId\":\"leaf\",\"lifecycle\":\"trashed\"}"), mapper.readTree(messages.getAllValues().get(0).getPayload()));
        org.junit.jupiter.api.Assertions.assertThrows(IllegalArgumentException.class, () -> registry.broadcastLifecycleChanged("leaf", "forged"));
    }

    @Test
    void revokedOrUnauthorizedSessionIsClosedBeforeReceivingRemoteRevision() throws Exception {
        AccessControlService access = mock(AccessControlService.class);
        WebSocketSessionRegistry registry = new WebSocketSessionRegistry(new ObjectMapper().findAndRegisterModules(), access,
                mock(RangeAccessService.class), mock(AccessProjectionService.class));
        WebSocketSession session = mock(WebSocketSession.class);
        Principal principal = authenticated("editor-1", Instant.now().plusSeconds(60));
        when(session.getAttributes()).thenReturn(new java.util.concurrent.ConcurrentHashMap<>(Map.of()));
        when(session.getPrincipal()).thenReturn(principal);
        when(session.isOpen()).thenReturn(true);
        doThrow(ServiceException.forbidden("Workbook access denied"))
                .when(access).require("book-1", "editor-1", WorkbookRole.VIEWER);
        registry.join("book-1", session);
        CommittedOperationEnvelope operation = new CommittedOperationEnvelope("test-session", 
                OperationEnvelope.SCHEMA,
                "operation-1",
                "book-1",
                "owner-1",
                OperationOrigin.CLIENT,
                1,
                0,
                1,
                List.of(new CommittedOperationMutation("cell.set", "sheet-1", new ObjectMapper().readTree("{\"row\":0,\"column\":0,\"value\":{\"value\":1}}"), List.of(new RangeRef("sheet-1", 0, 0, 0, 0)))),
                Instant.parse("2026-08-23T00:00:00Z"),
                Instant.parse("2026-08-23T00:00:00Z")
        );

        registry.broadcastRevision(operation);

        verify(session).close(eq(CloseStatus.POLICY_VIOLATION));
    }
    @Test
    void concurrentAccessAndCalculationBroadcastsShareOneConnectionWriter() throws Exception {
        var registry = new WebSocketSessionRegistry(new ObjectMapper(), mock(AccessControlService.class),
                mock(RangeAccessService.class), mock(AccessProjectionService.class));
        var session = mock(WebSocketSession.class);
        when(session.isOpen()).thenReturn(true);
        when(session.getPrincipal()).thenReturn(authenticated("reader", Instant.now().plusSeconds(60)));
        when(session.getAttributes()).thenReturn(new java.util.concurrent.ConcurrentHashMap<>());
        registry.join("root", session);
        registry.subscribeCalculation("root", java.util.Set.of("source"), session);
        var entered = new java.util.concurrent.CountDownLatch(1);
        var release = new java.util.concurrent.CountDownLatch(1);
        var writes = new java.util.concurrent.atomic.AtomicInteger();
        var active = new java.util.concurrent.atomic.AtomicInteger();
        org.mockito.Mockito.doAnswer(invocation -> {
            assertEquals(1, active.incrementAndGet(), "a WebSocket cannot accept concurrent writers");
            try {
                if (writes.incrementAndGet() == 1) {
                    entered.countDown();
                    if (!release.await(5, java.util.concurrent.TimeUnit.SECONDS)) throw new IllegalStateException("Writer timed out");
                }
            } finally { active.decrementAndGet(); }
            return null;
        }).when(session).sendMessage(any());
        try (var executor = java.util.concurrent.Executors.newFixedThreadPool(2)) {
            var first = executor.submit(() -> registry.broadcastAccessChanged("source", 7));
            org.junit.jupiter.api.Assertions.assertTrue(entered.await(5, java.util.concurrent.TimeUnit.SECONDS));
            var second = executor.submit(() -> registry.broadcastAccessChanged("root", 8));
            try {
                org.junit.jupiter.api.Assertions.assertThrows(java.util.concurrent.TimeoutException.class,
                        () -> second.get(100, java.util.concurrent.TimeUnit.MILLISECONDS));
            } finally { release.countDown(); }
            first.get(5, java.util.concurrent.TimeUnit.SECONDS);
            second.get(5, java.util.concurrent.TimeUnit.SECONDS);
        }
        assertEquals(2, writes.get());
        verify(session, never()).close(any());
    }

    @Test
    void failedTransportRetiresMembershipAndCalculationWithoutRejectingCommittedFacts() throws Exception {
        for (Exception failure : List.of(new java.io.IOException("Transport closed"), new IllegalStateException("Writer unavailable"))) {
            var registry = new WebSocketSessionRegistry(new ObjectMapper(), mock(AccessControlService.class),
                    mock(RangeAccessService.class), mock(AccessProjectionService.class));
            var session = mock(WebSocketSession.class);
            when(session.isOpen()).thenReturn(true);
            when(session.getPrincipal()).thenReturn(authenticated("reader", Instant.now().plusSeconds(60)));
            when(session.getAttributes()).thenReturn(new java.util.concurrent.ConcurrentHashMap<>());
            registry.join("root", session);
            registry.subscribeCalculation("root", java.util.Set.of("source"), session);
            doThrow(failure).when(session).sendMessage(any());
            org.junit.jupiter.api.Assertions.assertDoesNotThrow(() -> registry.broadcastAccessChanged("source", 7));
            verify(session).close(eq(CloseStatus.SERVER_ERROR));
            assertEquals(null, registry.unitId(session));
            registry.broadcastAccessChanged("source", 8);
            registry.broadcastAccessChanged("root", 9);
            verify(session, times(1)).sendMessage(any());
        }
    }

}
