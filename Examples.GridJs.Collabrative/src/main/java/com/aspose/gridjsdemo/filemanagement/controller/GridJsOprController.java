package com.aspose.gridjsdemo.filemanagement.controller;

import java.security.Principal;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.messaging.handler.annotation.DestinationVariable;
import org.springframework.messaging.handler.annotation.MessageMapping;
import org.springframework.messaging.handler.annotation.Payload;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.aspose.gridjs.GridJsCoService;
import com.aspose.gridjs.OprMessage;
import com.aspose.gridjs.OprMessageDto;

import lombok.RequiredArgsConstructor;
@Controller
@RequiredArgsConstructor
@RestController  
@RequestMapping("/GridJs2/msg")  
public class GridJsOprController {

	 
	
	@Autowired
private GridJsCoService gridJsCoService;
	
 
	


private final String historyCountString="5000";
  
// public MessageController(SimpMessagingTemplate messagingTemplate) {
//        this.messagingTemplate = messagingTemplate;
//    }
 /**
 * 处理客户端发送的消息 4. 最终建议 1如果吞吐量要求不高：直接使用 @Transactional，简单可靠。
 * 
 * 2如果需要高性能：
 * 
 * 优先选择 @TransactionalEventListener（平衡一致性和性能）。 -->OprMessageEventListener // 使用
 * SimpMessagingTemplate 广播消息
 * messagingTemplate.convertAndSend("/topic/opr/"+fileId, saved);
 * 
 * 3对于超高并发场景，引入 消息队列 + 本地事务表。
 * 
 * 
 * 
 */
@MessageMapping("/opr/{fileId}")
// @SendTo("/topic/opr/{fileId}")
public void handleMessage(@DestinationVariable String fileId, @Payload OprMessageDto dto, Principal principal) {
	 // 获取当前登录用户名
	
	  gridJsCoService.handleMessage(fileId, dto);

     }



@GetMapping("/history/{fileId}")
public ResponseEntity<List<OprMessage>> getHistory(@PathVariable String fileId,
		@RequestParam(required = false, defaultValue = historyCountString) int limit) {
    
	try {
		 
		gridJsCoService.checkAndSyncForCollaborative(fileId);
	} catch (Exception e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
	List<OprMessage> history = gridJsCoService.getFileHistory(fileId, limit);
    return ResponseEntity.ok(history);
}
}
